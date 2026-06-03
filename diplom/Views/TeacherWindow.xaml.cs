using diplom.Models;
using System;
using System.Linq;
using System.Windows;

namespace diplom.Views
{
    /// <summary>
    /// Логика взаимодействия для TeacherWindow.xaml
    /// </summary>
    public partial class TeacherWindow : Window
    {
        private int _teacherId;

        public TeacherWindow(int teacherId, string greeting)
        {
            InitializeComponent();
            _teacherId = teacherId;
            TxtGreeting.Text = greeting;
            LoadDisciplines();
            LoadDebts();
        }

        private void LoadDisciplines()
        {
            var discs = OpenContext.db.Disciplines
                .Where(d => d.Teachers.Any(t => t.TeacherID == _teacherId))
                .ToList();
            // Добавляем вариант для всех дисциплин
            var all = new Discipline { DisciplineID = 0, DisciplineName = "Все дисциплины" };
            discs.Insert(0, all);
            CbDisciplines.ItemsSource = discs;
            CbDisciplines.SelectedValue = 0;
        }

        private void LoadDebts(int? disciplineId = null)
        {
            var debts = OpenContext.db.Debts
                .Where(d => d.TeacherID == _teacherId && !d.IsCleared);

            if (disciplineId.HasValue)
                debts = debts.Where(d => d.DisciplineID == disciplineId.Value);

            var list = debts.Select(d => new
            {
                d.DebtID,
                StudentName = d.Student.LastName + " " + d.Student.FirstName,
                StudentCardNumber = d.Student.StudentCardNumber,
                DisciplineName = d.Discipline.DisciplineName,
                d.DebtStatus,
                DateRecorded = d.DateRecorded,
                d.IsCleared
            }).ToList();

            DgDebts.ItemsSource = list;
        }

        private void CbDisciplines_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (CbDisciplines.SelectedValue is int id && id != 0)
                LoadDebts(id);
            else
                LoadDebts();
        }

        private void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            LoadDisciplines();
            LoadDebts();
        }

        private void BtnMarkCleared_Click(object sender, RoutedEventArgs e)
        {
            var selected = DgDebts.SelectedItem;
            if (selected == null)
            {
                MessageBox.Show("Пожалуйста, выберите задолженность.", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                // Попробуем получить DebtID безопасно
                var prop = selected.GetType().GetProperty("DebtID");
                if (prop == null)
                {
                    MessageBox.Show("Не удалось определить идентификатор задолженности.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                int debtId = (int)prop.GetValue(selected, null);

                var debt = OpenContext.db.Debts.FirstOrDefault(d => d.DebtID == debtId);
                if (debt == null)
                {
                    MessageBox.Show("Задолженность не найдена.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                debt.IsCleared = true;
                debt.DateCleared = DateTime.Now;
                debt.DebtStatus = "Сдан";

                // Явно пометим сущность как изменённую на случай проблем с трекингом
                var entry = OpenContext.db.Entry(debt);
                entry.State = System.Data.Entity.EntityState.Modified;

                OpenContext.db.SaveChanges();

                MessageBox.Show("Задолженность отмечена как сданная.", "Готово", MessageBoxButton.OK, MessageBoxImage.Information);

                LoadDebts();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnExit_Click(object sender, RoutedEventArgs e)
        {
            MainWindow window = new MainWindow();
            window.Show();
            this.Close();
        }
    }
}
