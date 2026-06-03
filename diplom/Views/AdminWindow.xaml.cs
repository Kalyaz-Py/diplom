using diplom.Models;
using Microsoft.Win32;
using OfficeOpenXml;
using Xceed.Words.NET;
using System.Data.Entity;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Xceed.Document.NET;

namespace diplom.Views
{
    /// <summary>
    /// Логика взаимодействия для AdminWindow.xaml
    /// </summary>
    public partial class AdminWindow : Window
    {

        public AdminWindow()
        {
            InitializeComponent();
            LoadDebtsData();
        }

        // Экспорт в Excel в формате, соответствующем структуре Word-примера
        private void ExportDebtsToExcel(string filePath)
        {
            try
            {
                // Загружаем данные. Явно выгружаем справочники и связываем в памяти, чтобы избежать проблем с Include
                var debtsRaw = OpenContext.db.Debts
                    .Where(d => d.DebtStatus == "Активна")
                    .ToList();

                var students = OpenContext.db.Students.ToList();
                var groups = OpenContext.db.Groups.ToList();
                var disciplines = OpenContext.db.Disciplines.ToList();

                // Формируем объекты с подгруженными навигационными свойствами (явно, безопасно)
                var debts = new List<dynamic>();
                var skipped = new List<int>(); // IDs долгов, которые не удалось сопоставить
                foreach (var d in debtsRaw)
                {
                    var student = students.SingleOrDefault(s => s.StudentID == d.StudentID);
                    if (student == null) { skipped.Add(d.DebtID); continue; }
                    var group = groups.SingleOrDefault(g => g.GroupID == student.GroupID);
                    if (group == null) { skipped.Add(d.DebtID); continue; }
                    var discipline = disciplines.SingleOrDefault(di => di.DisciplineID == d.DisciplineID);
                    if (discipline == null) { skipped.Add(d.DebtID); continue; }

                    debts.Add(new { Debt = d, Student = student, Group = group, Discipline = discipline });
                }

                // Группируем по идентификатору группы, затем по имени (чтобы избежать слияния разных групп с одинаковыми именами)
                var byGroup = debts
                    .GroupBy(x => x.Group.GroupID)
                    .Select(g => new { Group = g.First().Group, Debts = g.ToList() })
                    .OrderBy(g => g.Group.GroupName)
                    .ToList();

                // Диагностика: подсчёт
                int totalRaw = debtsRaw.Count;
                int totalProcessed = debts.Count;
                int uniqueStudents = debts.Select(x => x.Student.StudentID).Distinct().Count();
                int uniqueGroups = debts.Select(x => x.Group.GroupID).Distinct().Count();

                // Создаём Excel
                using (var package = new OfficeOpenXml.ExcelPackage())
                {
                    // Создаём один лист со всеми группами и всеми должниками
                    var sheet = package.Workbook.Worksheets.Add("Ведомость задолженностей");

                    // Заголовки столбцов
                    sheet.Cells[1, 1].Value = "Код группы";
                    sheet.Cells[1, 2].Value = "ФИО студента";
                    sheet.Cells[1, 3].Value = "Задолженности";

                    int row = 2;

                    foreach (var group in byGroup)
                    {
                        var groupDebts = group.Debts;

                        var byStudent = groupDebts
                            .GroupBy(d => new { d.Student.StudentID, FullName = string.Join(" ", new[] { d.Student.LastName, d.Student.FirstName, d.Student.MiddleName }.Where(s => !string.IsNullOrWhiteSpace(s))) })
                            .OrderBy(sg => sg.Key.FullName)
                            .ToList();

                        bool firstStudentInGroup = true;
                        foreach (var studentGroup in byStudent)
                        {
                            var studentDebts = studentGroup.ToList();
                            if (studentDebts.Count == 0)
                                continue;

                            // Пишем название группы только для первого студента в группе, для остальных оставляем пустую ячейку
                            if (firstStudentInGroup)
                            {
                                sheet.Cells[row, 1].Value = group.Group.GroupName;
                                firstStudentInGroup = false;
                            }
                            else
                            {
                                sheet.Cells[row, 1].Value = string.Empty;
                            }
                            sheet.Cells[row, 2].Value = studentGroup.Key.FullName;

                            // Группируем долги по типу контроля
                            var byControl = studentDebts
                                .GroupBy(d => (d.Discipline?.ControlType) ?? "Не указан")
                                .OrderBy(g => g.Key)
                                .ToList();

                            // Формируем содержимое ячейки с RichText: название формы контроля — жирным, каждый предмет на новой строке
                            var cell = sheet.Cells[row, 3];
                            cell.Value = null; // сброс
                            cell.RichText.Clear();
                            bool firstControl = true;
                            foreach (var control in byControl)
                            {
                                var controlLabel = control.Key;
                                var subjects = control.Select(d => d.Discipline.DisciplineName).Where(n => !string.IsNullOrWhiteSpace(n)).Distinct().ToList();
                                if (subjects.Count == 0)
                                    continue;

                                // Добавляем название формы контроля жирным
                                if (!firstControl)
                                {
                                    // Отделяем предыдущий блок переводом строки
                                    cell.RichText.Add(Environment.NewLine);
                                }
                                var rt = cell.RichText.Add(controlLabel + ":");
                                rt.Bold = true;

                                // Добавляем предметы в виде нумерованного списка, каждый с новой строки
                                for (int i = 0; i < subjects.Count; i++)
                                {
                                    var subjText = Environment.NewLine + ($"{i + 1}.\t{subjects[i]}");
                                    var r = cell.RichText.Add(subjText);
                                    r.Bold = false;
                                }

                                firstControl = false;
                            }

                            cell.Style.WrapText = true;
                            row++;
                        }
                    }

                    // Автоподгонка колонок по использованной области
                    if (sheet.Dimension != null)
                        sheet.Cells[sheet.Dimension.Address].AutoFitColumns();

                    // Дополнительный лист: студенты в группе риска (более 3 активных долгов)
                    try
                    {
                        var riskList = (from d in debtsRaw
                                        join s in students on d.StudentID equals s.StudentID
                                        join g in groups on s.GroupID equals g.GroupID
                                        where !(d.IsCleared == true || d.DebtStatus == "Сдан")
                                        group d by new { s.StudentID, s.StudentCardNumber, s.LastName, s.FirstName, s.MiddleName, g.GroupName } into grp
                                        let cnt = grp.Count()
                                        where cnt >= 3
                                        select new
                                        {
                                            GroupName = grp.Key.GroupName,
                                            StudentName = string.Join(" ", new[] { grp.Key.LastName, grp.Key.FirstName, grp.Key.MiddleName }.Where(x => !string.IsNullOrWhiteSpace(x))),
                                            StudentCardNumber = grp.Key.StudentCardNumber,
                                            DebtCount = cnt
                                        }).OrderByDescending(x => x.DebtCount).ThenBy(x => x.GroupName).ToList();

                        if (riskList != null && riskList.Count > 0)
                        {
                            var riskSheet = package.Workbook.Worksheets.Add("Группа риска");
                            riskSheet.Cells[1, 1].Value = "Группа";
                            riskSheet.Cells[1, 2].Value = "ФИО студента";
                            riskSheet.Cells[1, 3].Value = "№ зачетки";
                            riskSheet.Cells[1, 4].Value = "Количество долгов";

                            for (int col = 1; col <= 4; col++)
                            {
                                var cell = riskSheet.Cells[1, col];
                                cell.Style.Font.Bold = true;
                                cell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            }

                            int r = 2;
                            foreach (var item in riskList)
                            {
                                riskSheet.Cells[r, 1].Value = item.GroupName;
                                riskSheet.Cells[r, 2].Value = item.StudentName;
                                riskSheet.Cells[r, 3].Value = item.StudentCardNumber;
                                riskSheet.Cells[r, 4].Value = item.DebtCount;
                                r++;
                            }

                            if (riskSheet.Dimension != null)
                                riskSheet.Cells[riskSheet.Dimension.Address].AutoFitColumns();
                        }
                    }
                    catch { }

                    // Сохраняем в файл
                    var fi = new FileInfo(filePath);
                    package.SaveAs(fi);
                }

                // Если были пропуски - сохраним простую диагностику в файл рядом с экспортом
                if (skipped.Count > 0)
                {
                    try
                    {
                        var diagPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(filePath), System.IO.Path.GetFileNameWithoutExtension(filePath) + "_skipped.csv");
                        File.WriteAllLines(diagPath, skipped.Select(id => id.ToString()));
                        MessageBox.Show($"Экспорт завершён: {filePath}\nОбработано записей: {totalProcessed} из {totalRaw}. Уникальных студентов: {uniqueStudents}. Уникальных групп: {uniqueGroups}.\nНекоторые записи были пропущены (не найдены связанные сущности). Список ID долгов сохранён: {diagPath}", "Экспорт", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch
                    {
                        MessageBox.Show($"Экспорт завершён: {filePath}\nОбработано записей: {totalProcessed} из {totalRaw}. Уникальных студентов: {uniqueStudents}. Уникальных групп: {uniqueGroups}.\nНекоторые записи были пропущены (не найдены связанные сущности).", "Экспорт", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                else
                {
                    MessageBox.Show($"Экспорт в Excel завершён: {filePath}\nОбработано записей: {totalProcessed} из {totalRaw}. Уникальных студентов: {uniqueStudents}. Уникальных групп: {uniqueGroups}.", "Экспорт", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта в Excel: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            // Открыть диалог сохранения файла .docx и вызвать экспорт ведомости задолженностей
            try
            {
                var dlg = new Microsoft.Win32.SaveFileDialog();
                dlg.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
                dlg.FileName = "Ведомость_задолженностей.xlsx";
                if (dlg.ShowDialog() == true)
                {
                    ExportDebtsToExcel(dlg.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Метод для загрузки и отображения всех активных долгов в DataGrid
        private void LoadDebtsData()
        {
            try
            {
                var debtsList = (from d in OpenContext.db.Debts
                                 join s in OpenContext.db.Students on d.StudentID equals s.StudentID
                                 join g in OpenContext.db.Groups on s.GroupID equals g.GroupID
                                 join disc in OpenContext.db.Disciplines on d.DisciplineID equals disc.DisciplineID
                                 join t in OpenContext.db.Teachers on d.TeacherID equals t.TeacherID
                                 join sched in OpenContext.db.Schedules on d.DebtID equals sched.DebtID into schedules
                                 from sched in schedules.DefaultIfEmpty()
                                 join c in OpenContext.db.Classrooms on (sched != null ? sched.ClassroomID : (int?)null) equals c.ClassroomID into classrooms
                                 from c in classrooms.DefaultIfEmpty()
                                 where d.DebtStatus == "Активна"
                                 select new
                                 {
                                     StudentCardNumber = s.StudentCardNumber,
                                     StudentName = s.LastName + " " + s.FirstName + " " + s.MiddleName,
                                     GroupName = g.GroupName,
                                     DisciplineName = disc.DisciplineName,
                                     TeacherName = t.LastName + " " + t.FirstName.Substring(0, 1) + "." + (t.MiddleName != null ? t.MiddleName.Substring(0, 1) + "." : ""),
                                     Status = d.DebtStatus,
                                     TimeSlotID = sched != null ? (int?)sched.TimeSlotID : (int?)null,
                                     ClassroomNumber = c != null ? c.ClassroomNumber : "Не назначен",
                                     ExamDate = sched != null ? sched.ExamDate : (DateTime?)null,

                                 }).ToList();

                DgAdminDebts.ItemsSource = debtsList;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnAddDebt_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                AddDebtWindow addDebtWindow = new AddDebtWindow();
                bool? result = null;
                try
                {
                    result = addDebtWindow.ShowDialog();
                }
                catch (Exception dlgEx)
                {
                    // Если при открытии/выполнение окна возникла ошибка - показываем подробности
                    MessageBox.Show($"Ошибка при открытии окна добавления должника: {dlgEx.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Обновляем таблицу только если добавление прошло успешно (DialogResult == true)
                if (result == true)
                {
                    LoadDebtsData();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Необработанная ошибка при запуске окна добавления: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnImportExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Файлы Excel (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";

            if (openFileDialog.ShowDialog() == true)
            {
                FileInfo fileInfo = new FileInfo(openFileDialog.FileName);

                try
                {
                    TxtStatus.Text = "Статус: выполняется импорт...";

                    using (ExcelPackage package = new ExcelPackage(fileInfo))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                        int rowCount = worksheet.Dimension.Rows;

                        // Кешируем справочники для быстрого поиска
                        // Используем GroupBy -> First(), чтобы избежать ошибки при дубликатах ключей
                        var studentsByCard = OpenContext.db.Students
                            .ToList()
                            .GroupBy(s => (s.StudentCardNumber ?? string.Empty).Trim(), StringComparer.InvariantCultureIgnoreCase)
                            .ToDictionary(g => g.Key, g => g.First(), StringComparer.InvariantCultureIgnoreCase);

                        var disciplinesByName = OpenContext.db.Disciplines
                            .ToList()
                            .GroupBy(d => (d.DisciplineName ?? string.Empty).Trim(), StringComparer.InvariantCultureIgnoreCase)
                            .ToDictionary(g => g.Key, g => g.First(), StringComparer.InvariantCultureIgnoreCase);
                        var teachers = OpenContext.db.Teachers.ToList();

                        var report = new List<diplom.ViewModels.ImportReportItem>();
                        int importedCount = 0;
                        int batchCounter = 0;

                        // Определяем заголовки (если есть) — попробовать автоматически найти индексы колонок
                        int cardCol = 1, discCol = 2, gradeCol = 3;
                        var header = worksheet.Cells[1, 1, 1, worksheet.Dimension.Columns];
                        if (header != null)
                        {
                            for (int c = 1; c <= worksheet.Dimension.Columns; c++)
                            {
                                var val = (worksheet.Cells[1, c].Value ?? string.Empty).ToString().ToLower();
                                if (val.Contains("зачет") || val.Contains("зачетка") || val.Contains("card")) cardCol = c;
                                if (val.Contains("дисцип") || val.Contains("предмет") || val.Contains("discip")) discCol = c;
                                if (val.Contains("оценк") || val.Contains("grade") || val.Contains("mark")) gradeCol = c;
                            }
                        }

                        for (int row = 2; row <= rowCount; row++)
                        {
                            var cardValue = worksheet.Cells[row, cardCol].Value;
                            var disciplineValue = worksheet.Cells[row, discCol].Value;
                            var gradeValue = worksheet.Cells[row, gradeCol].Value;

                            string cardNumber = cardValue?.ToString().Trim();
                            string disciplineName = disciplineValue?.ToString().Trim();
                            string grade = gradeValue?.ToString().Trim();

                            if (string.IsNullOrWhiteSpace(cardNumber) && string.IsNullOrWhiteSpace(disciplineName) && string.IsNullOrWhiteSpace(grade))
                                continue; // пустая строка

                            // Нормализация
                            cardNumber = cardNumber ?? string.Empty;
                            disciplineName = disciplineName ?? string.Empty;
                            grade = grade ?? string.Empty;

                            // Определяем маркер неудов
                            bool isDebt = false;
                            var low = grade.ToLower();
                            if (low == "2" || low.Contains("не зачт") || low.Contains("н/я") || low.Contains("ня") || low.Contains("неуд"))
                                isDebt = true;

                            if (!isDebt)
                            {
                                report.Add(new diplom.ViewModels.ImportReportItem { Row = row, CardNumber = cardNumber, Discipline = disciplineName, Grade = grade, Reason = "Оценка не помечает долг" });
                                continue;
                            }

                            // Ищем в кеше
                            studentsByCard.TryGetValue(cardNumber, out var student);
                            disciplinesByName.TryGetValue(disciplineName, out var discipline);

                            if (student == null)
                            {
                                report.Add(new diplom.ViewModels.ImportReportItem { Row = row, CardNumber = cardNumber, Discipline = disciplineName, Grade = grade, Reason = "Студент не найден" });
                                continue;
                            }

                            if (discipline == null)
                            {
                                report.Add(new diplom.ViewModels.ImportReportItem { Row = row, CardNumber = cardNumber, Discipline = disciplineName, Grade = grade, Reason = "Дисциплина не найдена" });
                                continue;
                            }

                            // Проверяем дубликат
                            bool debtExists = OpenContext.db.Debts.Any(d => d.StudentID == student.StudentID && d.DisciplineID == discipline.DisciplineID && d.DebtStatus == "Активна");
                            if (debtExists)
                            {
                                report.Add(new diplom.ViewModels.ImportReportItem { Row = row, CardNumber = cardNumber, Discipline = disciplineName, Grade = grade, Reason = "Дубликат (активная задолженность уже есть)" });
                                continue;
                            }

                            // Создаём новую задолженность
                            var teacher = teachers.FirstOrDefault();
                            Debt newDebt = new Debt
                            {
                                StudentID = student.StudentID,
                                DisciplineID = discipline.DisciplineID,
                                TeacherID = teacher != null ? teacher.TeacherID : 1,
                                DebtStatus = "Активна",
                                DateRecorded = DateTime.Now
                            };

                            OpenContext.db.Debts.Add(newDebt);
                            importedCount++;
                            batchCounter++;

                            if (batchCounter >= 500)
                            {
                                OpenContext.db.SaveChanges();
                                batchCounter = 0;
                            }
                        }

                        if (batchCounter > 0) OpenContext.db.SaveChanges();

                        // Показать отчёт о непомещённых/пропущенных строках
                        if (report.Count > 0)
                        {
                            var reportWin = new diplom.Views.ImportReportWindow(report);
                            reportWin.Owner = this;
                            reportWin.ShowDialog();
                        }

                        MessageBox.Show($"Импорт завершен. Добавлено новых задолженностей: {importedCount}. См. отчет по пропускам.", "Импорт", MessageBoxButton.OK, MessageBoxImage.Information);
                        TxtStatus.Text = "Статус: импорт завершен";
                        LoadDebtsData();
                    }
                }
                catch (Exception ex)
                {
                    TxtStatus.Text = "Статус: ошибка при импорте";
                    MessageBox.Show($"Ошибка при чтении файла Excel: {ex.Message}", "Критическая ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void BtnRunAlgorithm_Click(object sender, RoutedEventArgs e)
        {
            if (!OpenContext.db.Debts.Any(d => d.DebtStatus == "Активна"))
            {
                MessageBox.Show("В системе нет активных задолженностей для формирования расписания.", "Уведомление", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // Открываем окно для выбора параметров дат
            ScheduleGeneratorWindow scheduleWindow = new ScheduleGeneratorWindow();
            if (scheduleWindow.ShowDialog() == true)
            {
                try
                {
                    TxtStatus.Text = "Статус: выполняется формирование расписания...";

                    // Собираем список выбранных студентов и инициализируем планировщик с параметрами дат
                    var selectedStudentIds = scheduleWindow.SelectedStudentIds;
                    GeneticScheduler scheduler = new GeneticScheduler(scheduleWindow.ScheduleStartDate.Value, scheduleWindow.ScheduleEndDate.Value, selectedStudentIds);
                    // Запуск эволюционного отбора решений
                    var optimalChromosome = scheduler.RunEvolution();

                    if (optimalChromosome != null)
                    {
                        // Удаляем только те записи старого расписания, которые попадают в выбранный период
                        // Если выбраны конкретные студенты - удаляем только их записи в периоде, иначе удаляем все записи в периоде
                        var start = scheduleWindow.ScheduleStartDate;
                        var end = scheduleWindow.ScheduleEndDate;
                        DateTime startVal = DateTime.MinValue;
                        DateTime endVal = DateTime.MaxValue;
                        if (start.HasValue && end.HasValue)
                        {
                            startVal = start.Value.Date;
                            endVal = end.Value.Date;

                            if (selectedStudentIds != null && selectedStudentIds.Count > 0)
                            {
                                // получаем долги для выбранных студентов
                                var debtIds = OpenContext.db.Debts
                                    .Where(d => selectedStudentIds.Contains(d.StudentID))
                                    .Select(d => d.DebtID)
                                    .ToList();

                                var oldSchedules = OpenContext.db.Schedules
                                    .Where(s => s.ExamDate >= startVal && s.ExamDate <= endVal
                                                && debtIds.Contains(s.DebtID))
                                    .ToList();

                                OpenContext.db.Schedules.RemoveRange(oldSchedules);
                            }
                            else
                            {
                                var oldSchedules = OpenContext.db.Schedules
                                    .Where(s => s.ExamDate >= startVal && s.ExamDate <= endVal)
                                    .ToList();

                                OpenContext.db.Schedules.RemoveRange(oldSchedules);
                            }

                            OpenContext.db.SaveChanges();
                        }

                    // Переносим гены победившей хромосомы в таблицы базы данных SQL Server
                        // Фильтруем по выбранным датам и при необходимости по выбранным студентам
                        int addedSchedules = 0;
                    var updatedList = new System.Collections.Generic.List<diplom.ViewModels.UpdatedScheduleInfo>();
                        foreach (var gene in optimalChromosome.Genes)
                        {
                            // Проверяем, что дата экзамена попадает в выбранный диапазон
                            if (gene.ExamDate < startVal || gene.ExamDate > endVal)
                                continue;

                            // Если пользователь выбрал конкретных студентов, проверяем принадлежность долга
                            if (selectedStudentIds != null && selectedStudentIds.Count > 0)
                            {
                                var debt = OpenContext.db.Debts.FirstOrDefault(d => d.DebtID == gene.DebtID);
                                if (debt == null || !selectedStudentIds.Contains(debt.StudentID))
                                    continue; // пропускаем, не в выбранном списке
                            }

                            // Если для этого долга уже есть запись расписания - обновляем её, иначе создаем новую
                            var existingSchedule = OpenContext.db.Schedules.FirstOrDefault(s => s.DebtID == gene.DebtID);
                            if (existingSchedule != null)
                            {
                                existingSchedule.ExamDate = gene.ExamDate;
                                existingSchedule.TimeSlot = gene.TimeSlot;
                                existingSchedule.ClassroomID = gene.ClassroomID;

                                // Подготовим запись TimeSlot для отображения и сохранения FK
                                TimeSlot tsRec = null;
                                if (gene.TimeSlotID > 0)
                                {
                                    tsRec = OpenContext.db.TimeSlots.FirstOrDefault(ts => ts.TimeSlotID == gene.TimeSlotID);
                                    if (tsRec != null)
                                        existingSchedule.TimeSlotID = tsRec.TimeSlotID;
                                    else
                                        existingSchedule.TimeSlotID = gene.TimeSlotID; // fallback
                                }
                                else
                                {
                                    tsRec = OpenContext.db.TimeSlots.FirstOrDefault(ts => ts.SlotNumber == gene.TimeSlot);
                                    if (tsRec != null)
                                    {
                                        existingSchedule.TimeSlotID = tsRec.TimeSlotID;
                                    }
                                }

                                // Сохраняем информацию об обновлённой записи для показа
                                try
                                {
                                    var debt = OpenContext.db.Debts.FirstOrDefault(d => d.DebtID == gene.DebtID);
                                    var student = debt != null ? OpenContext.db.Students.FirstOrDefault(s => s.StudentID == debt.StudentID) : null;
                                    var group = student != null ? OpenContext.db.Groups.FirstOrDefault(g => g.GroupID == student.GroupID) : null;
                                    var discipline = debt != null ? OpenContext.db.Disciplines.FirstOrDefault(di => di.DisciplineID == debt.DisciplineID) : null;
                                    var teacher = debt != null ? OpenContext.db.Teachers.FirstOrDefault(te => te.TeacherID == debt.TeacherID) : null;
                                    var classroom = OpenContext.db.Classrooms.FirstOrDefault(cl => cl.ClassroomID == existingSchedule.ClassroomID);

                                    updatedList.Add(new diplom.ViewModels.UpdatedScheduleInfo
                                    {
                                        DebtID = gene.DebtID,
                                        StudentCardNumber = student != null ? student.StudentCardNumber : string.Empty,
                                        StudentName = student != null ? student.LastName + " " + student.FirstName + " " + student.MiddleName : string.Empty,
                                        GroupName = group != null ? group.GroupName : string.Empty,
                                        DisciplineName = discipline != null ? discipline.DisciplineName : string.Empty,
                                        TeacherName = teacher != null ? teacher.LastName + " " + teacher.FirstName.Substring(0,1) + "." : string.Empty,
                                        Status = debt != null ? debt.DebtStatus : string.Empty,
                                        ExamDate = existingSchedule.ExamDate,
                                        TimeSlotNumber = gene.TimeSlot,
                                        TimeSlotDescription = tsRec != null ? tsRec.Description : ("Пара " + gene.TimeSlot.ToString()),
                                        ClassroomNumber = classroom != null ? classroom.ClassroomNumber : "Не назначен",
                                        IsNew = false
                                    });
                                }
                                catch { }
                            }
                            else
                            {
                                Schedule dbRecord = new Schedule
                                {
                                    DebtID = gene.DebtID,
                                    ExamDate = gene.ExamDate,
                                    TimeSlot = gene.TimeSlot,
                                    ClassroomID = gene.ClassroomID
                                };
                                // Сохраняем TimeSlotID напрямую, если он присутствует и подготовим tsRecNew для отображения
                                TimeSlot tsRecNew = null;
                                if (gene.TimeSlotID > 0)
                                {
                                    tsRecNew = OpenContext.db.TimeSlots.FirstOrDefault(ts => ts.TimeSlotID == gene.TimeSlotID);
                                    if (tsRecNew != null)
                                        dbRecord.TimeSlotID = tsRecNew.TimeSlotID;
                                    else
                                        dbRecord.TimeSlotID = gene.TimeSlotID; // fallback
                                }
                                else
                                {
                                    tsRecNew = OpenContext.db.TimeSlots.FirstOrDefault(ts => ts.SlotNumber == gene.TimeSlot);
                                    if (tsRecNew != null)
                                    {
                                        dbRecord.TimeSlotID = tsRecNew.TimeSlotID;
                                    }
                                }

                                OpenContext.db.Schedules.Add(dbRecord);
                                addedSchedules++;

                                // Собираем информацию о новой записи
                                try
                                {
                                    var debt = OpenContext.db.Debts.FirstOrDefault(d => d.DebtID == gene.DebtID);
                                    var student = debt != null ? OpenContext.db.Students.FirstOrDefault(s => s.StudentID == debt.StudentID) : null;
                                    var group = student != null ? OpenContext.db.Groups.FirstOrDefault(g => g.GroupID == student.GroupID) : null;
                                    var discipline = debt != null ? OpenContext.db.Disciplines.FirstOrDefault(di => di.DisciplineID == debt.DisciplineID) : null;
                                    var teacher = debt != null ? OpenContext.db.Teachers.FirstOrDefault(te => te.TeacherID == debt.TeacherID) : null;
                                    var classroom = OpenContext.db.Classrooms.FirstOrDefault(cl => cl.ClassroomID == dbRecord.ClassroomID);

                                    updatedList.Add(new diplom.ViewModels.UpdatedScheduleInfo
                                    {
                                        DebtID = gene.DebtID,
                                        StudentCardNumber = student != null ? student.StudentCardNumber : string.Empty,
                                        StudentName = student != null ? student.LastName + " " + student.FirstName + " " + student.MiddleName : string.Empty,
                                        GroupName = group != null ? group.GroupName : string.Empty,
                                        DisciplineName = discipline != null ? discipline.DisciplineName : string.Empty,
                                        TeacherName = teacher != null ? teacher.LastName + " " + teacher.FirstName.Substring(0,1) + "." : string.Empty,
                                        Status = debt != null ? debt.DebtStatus : string.Empty,
                                        ExamDate = dbRecord.ExamDate,
                                        TimeSlotNumber = gene.TimeSlot,
                                        TimeSlotDescription = tsRecNew != null ? tsRecNew.Description : ("Пара " + gene.TimeSlot.ToString()),
                                        ClassroomNumber = classroom != null ? classroom.ClassroomNumber : "Не назначен",
                                        IsNew = true
                                    });
                                }
                                catch { }
                            }
                        }

                        // Фиксируем транзакцию в СУБД
                        OpenContext.db.SaveChanges();

                        // Если есть детальные конфликты — сохраним их в CSV-файл рядом с приложением (для анализа)
                        try
                        {
                            if (optimalChromosome.ConflictDetails != null && optimalChromosome.ConflictDetails.Count > 0)
                            {
                                var reportPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "genetic_conflicts_report.csv");
                                System.IO.File.WriteAllLines(reportPath, optimalChromosome.ConflictDetails);
                                // Сообщим пользователю путь к файлу (тихое уведомление)
                                TxtStatus.Text = "Статус: отчет конфликтов сохранен в genetic_conflicts_report.csv";

                                try
                                {
                                    var win = new diplom.Views.ConflictReportWindow(optimalChromosome.ConflictDetails, reportPath);
                                    win.Owner = this;
                                    win.ShowDialog();
                                }
                                catch { }
                            }
                        }
                        catch { }

                        if (updatedList != null && updatedList.Count > 0)
                        {
                            try
                            {
                                diplom.Views.UpdatedSchedulesWindow win = new UpdatedSchedulesWindow(updatedList);
                                win.Owner = this;
                                win.ShowDialog();
                            }
                            catch { }
                        }

                        MessageBox.Show($"Оптимальное расписание успешно сгенерировано!\n" +
                                        $"Период: {scheduleWindow.ScheduleStartDate:dd.MM.yyyy} - {scheduleWindow.ScheduleEndDate:dd.MM.yyyy}\n" +
                                        $"Добавлено экзаменов: {addedSchedules}\n" +
                                        $"Накладок/конфликтов обнаружено: {optimalChromosome.FitnessScore}\n" +
                                        $"Данные сохранены в БД.",
                                        "Успех эволюционного отбора", MessageBoxButton.OK, MessageBoxImage.Information);

                        TxtStatus.Text = "Статус: расписание успешно сформировано";

                        // Обновляем таблицу задолженностей на экране
                        LoadDebtsData();
                    }
                }
                catch (Exception ex)
                {
                    TxtStatus.Text = "Статус: ошибка при формировании расписания";
                    MessageBox.Show($"Критическая ошибка при работе ИИ-модуля: {ex.Message}", "Ошибка алгоритма", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void BtnViewDebts_Click(object sender, RoutedEventArgs e)
        {
            DebtsFilterWindow debtsWindow = new DebtsFilterWindow();
            debtsWindow.ShowDialog();
        }

        private void BtnManageAssignments_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var win = new ManageAssignmentsWindow();
                win.Owner = this;
                win.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при открытии окна привязки преподавателей: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }



        private void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                TxtStatus.Text = "Статус: выполняется обновление...";

                // 1. Сбрасываем локальный кэш Entity Framework, чтобы прочитать чистые данные из SQL Server
                foreach (var entry in OpenContext.db.ChangeTracker.Entries())
                {
                    entry.Reload();
                }

                // 2. Ищем и удаляем дубликаты активных задолженностей
                // Группируем долги по студенту, предмету, преподавателю и статусу "Активна"
                var duplicateGroups = OpenContext.db.Debts
                    .Where(d => d.DebtStatus == "Активна")
                    .GroupBy(d => new { d.StudentID, d.DisciplineID, d.TeacherID })
                    .Where(g => g.Count() > 1) // Нас интересуют только те группы, где записей больше, чем 1
                    .ToList();

                int removedDuplicatesCount = 0;

                if (duplicateGroups.Count > 0)
                {
                    foreach (var group in duplicateGroups)
                    {
                        // Пропускаем самую первую (оригинальную) запись в группе
                        var duplicatesToRemove = group.Skip(1).ToList();

                        // Считаем, сколько дублей сейчас удалим
                        removedDuplicatesCount += duplicatesToRemove.Count;

                        // Удаляем дубликаты из контекста БД
                        OpenContext.db.Debts.RemoveRange(duplicatesToRemove);
                    }

                    // Сохраняем изменения в Microsoft SQL Server
                    OpenContext.db.SaveChanges();
                }

                // 3. Вызываем стандартную функцию перевыборки данных, чтобы обновить DataGrid на экране
                LoadDebtsData();

                // 4. Оповещаем пользователя, если дубликаты действительно были найдены и удалены
                if (removedDuplicatesCount > 0)
                {
                    MessageBox.Show($"База данных успешно оптимизирована!\nНайдено и удалено избыточных дубликатов: {removedDuplicatesCount}",
                                    "Очистка данных", MessageBoxButton.OK, MessageBoxImage.Information);
                    TxtStatus.Text = $"Статус: удалено {removedDuplicatesCount} дубликатов";
                }
                else
                {
                    TxtStatus.Text = "Статус: дубликаты не найдены";
                }
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "Статус: ошибка при обновлении";
                MessageBox.Show($"Ошибка при обновлении и очистке данных: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnBackToMain_Click(object sender, RoutedEventArgs e)
        {
            MainWindow mainWindow = new MainWindow();
            mainWindow.Show();
            this.Close();
        }

       
    }
}
