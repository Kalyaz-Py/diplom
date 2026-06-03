using diplom.ViewModels;
using diplom.Models;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Linq;

namespace diplom.Views
{
    public partial class DebtsFilterWindow : Window
    {
        private DebtsFilterViewModel _viewModel;

        public DebtsFilterWindow()
        {
            InitializeComponent();
            _viewModel = new DebtsFilterViewModel();
            this.DataContext = _viewModel;
            _viewModel.PropertyChanged += ViewModel_PropertyChanged;
            UpdateRecordCount();
        }

        private void ViewModel_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(_viewModel.FilteredDebts))
            {
                UpdateRecordCount();
            }
        }

        private void BtnExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Файлы Excel (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";
            saveFileDialog.FileName = $"Долги_студентов_{DateTime.Now:dd_MM_yyyy}";

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    FileInfo fileInfo = new FileInfo(saveFileDialog.FileName);

                    using (ExcelPackage package = new ExcelPackage())
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Долги");

                        // Заголовки колонок
                        worksheet.Cells[1, 1].Value = "№ зачетки";
                        worksheet.Cells[1, 2].Value = "Студент (ФИО)";
                        worksheet.Cells[1, 3].Value = "Группа";
                        worksheet.Cells[1, 4].Value = "Дисциплина";
                        worksheet.Cells[1, 5].Value = "Преподаватель";
                        worksheet.Cells[1, 6].Value = "Кабинет";
                        worksheet.Cells[1, 7].Value = "Дата экзамена";
                        worksheet.Cells[1, 8].Value = "Статус";

                        // Стилизация заголовков
                        for (int col = 1; col <= 8; col++)
                        {
                            var cell = worksheet.Cells[1, col];
                            cell.Style.Font.Bold = true;
                            cell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        }

                        // Заполняем данные
                        int row = 2;
                        foreach (var debt in _viewModel.FilteredDebts)
                        {
                            worksheet.Cells[row, 1].Value = debt.StudentCardNumber;
                            worksheet.Cells[row, 2].Value = debt.StudentName;
                            worksheet.Cells[row, 3].Value = debt.GroupName;
                            worksheet.Cells[row, 4].Value = debt.DisciplineName;
                            worksheet.Cells[row, 5].Value = debt.TeacherName;
                            worksheet.Cells[row, 6].Value = debt.ClassroomNumber;
                            worksheet.Cells[row, 7].Value = debt.ExamDate.HasValue ? debt.ExamDate.Value.ToString("dd.MM.yyyy") : "Не назначена";
                            worksheet.Cells[row, 8].Value = debt.Status;

                            row++;
                        }

                        // Автоширина колонок
                        for (int col = 1; col <= 8; col++)
                        {
                            worksheet.Column(col).AutoFit(10, 50);
                        }

                        // Сохраняем файл
                        package.SaveAs(fileInfo);

                        MessageBox.Show($"Файл успешно сохранен!\nПуть: {saveFileDialog.FileName}",
                                        "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при экспорте в Excel: {ex.Message}", "Ошибка", 
                                    MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void UpdateRecordCount()
        {
            if (_viewModel?.FilteredDebts != null)
            {
                TxtRecordCount.Text = _viewModel.FilteredDebts.Count.ToString();
            }
        }

        private void BtnMarkPassed_Click(object sender, RoutedEventArgs e)
        {
            var selected = DgDebts.SelectedItem as ViewModels.DebtInfo;
            if (selected == null)
            {
                MessageBox.Show("Выберите запись долга в таблице.", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var confirm = MessageBox.Show($"Отметить долг студента {selected.StudentName} по {selected.DisciplineName} как сданный?", "Подтвердите", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (confirm != MessageBoxResult.Yes)
                return;

            try
            {
                var debt = OpenContext.db.Debts.FirstOrDefault(d => d.DebtID == selected.DebtID);
                if (debt == null)
                {
                    MessageBox.Show("Запись долга не найдена в базе.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Помечаем как сдан и сохраняем время сдачи
                debt.DebtStatus = "Сдан";
                try
                {
                    debt.IsCleared = true;
                    debt.DateCleared = DateTime.Now;
                }
                catch { /* ignore */ }

                // Удаляем связанные расписания для этого долга (если есть)
                var schedules = OpenContext.db.Schedules.Where(s => s.DebtID == debt.DebtID).ToList();
                foreach (var s in schedules)
                {
                    OpenContext.db.Schedules.Remove(s);
                }

                // Явно помечаем сущность как изменённую и сохраняем
                try
                {
                    var entry = OpenContext.db.Entry(debt);
                    entry.State = System.Data.Entity.EntityState.Modified;
                }
                catch { /* ignore */ }

                OpenContext.db.SaveChanges();

                MessageBox.Show("Долг отмечен как сдан.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);

                // Обновляем ViewModel и счётчик
                _viewModel.Refresh();
                UpdateRecordCount();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при изменении статуса: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


    }
}
