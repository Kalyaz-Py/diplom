using diplom.Views;
using diplom.Models;
using System;
using System.Collections.Generic;
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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace diplom
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        // Вход для студентов (без пароля)
        private void BtnStudentMode_Click(object sender, RoutedEventArgs e)
        {
            // Создаем экземпляр гостевого окна студента
            StudentGuestWindow studentWindow = new StudentGuestWindow();
            studentWindow.Show();

            // Закрываем текущее стартовое окно
            this.Close();
        }

        // Вход для Администратора / Заведующего отделением
        private void BtnAdminLogin_Click(object sender, RoutedEventArgs e)
        {
            string login = TxtLogin.Text.Trim();
            string password = TxtPassword.Password.Trim();

            // Проверка на заполненность полей
            if (string.IsNullOrEmpty(login) || string.IsNullOrEmpty(password))
            {
                MessageBox.Show("Пожалуйста, заполните все поля для авторизации.", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                // Ищем пользователя в таблице Users по логину и паролю
                var user = OpenContext.db.Users.FirstOrDefault(u => u.UserLogin == login && u.UserPassword == password);

                if (user != null)
                {
                    // Формируем приветственное сообщение в зависимости от роли
                    string greeting;

                    switch (user.RoleID)
                    {
                        case 3: // преподаватель
                            // Пытаемся найти связанную запись преподавателя
                            var teacher = OpenContext.db.Teachers.FirstOrDefault(t => t.UserID == user.UserID);
                            if (teacher != null)
                            {
                                string middle = string.IsNullOrWhiteSpace(teacher.MiddleName) ? "" : " " + teacher.MiddleName;
                                string fio = $"{teacher.LastName} {teacher.FirstName}{middle}".Trim();
                                greeting = $"Добро пожаловать, {fio}";
                            }
                            else
                            {
                                greeting = "Добро пожаловать, преподаватель";
                            }
                            break;
                        case 2: // заведующий отделением
                            greeting = "Добро пожаловать, заведующий отделением";
                            break;
                        case 1: // администратор
                        default:
                            greeting = "Добро пожаловать, администратор";
                            break;
                    }

                    MessageBox.Show(greeting, "Авторизация", MessageBoxButton.OK, MessageBoxImage.Information);

                    // Открываем соответствующее окно в зависимости от роли
                    if (user.RoleID == 3)
                    {
                        var teacher = OpenContext.db.Teachers.FirstOrDefault(t => t.UserID == user.UserID);
                        int teacherId = teacher != null ? teacher.TeacherID : 0;
                        TeacherWindow tw = new TeacherWindow(teacherId, greeting);
                        tw.Show();
                    }
                    else
                    {
                        AdminWindow adminWindow = new AdminWindow();
                        adminWindow.Show();
                    }

                    // Закрываем окно входа
                    this.Close();
                }
                else
                {
                    // Если учетные данные неверны — выводим предупреждение и очищаем пароль
                    MessageBox.Show("Неверный логин или пароль! Доступ заблокирован.", "Ошибка авторизации", MessageBoxButton.OK, MessageBoxImage.Error);
                    TxtPassword.Clear();
                }
            }
            catch (Exception ex)
            {
                // На случай ошибок подключения/запроса к БД показываем сообщение
                MessageBox.Show($"Ошибка при обращении к базе данных: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Просмотр долгов студентов (без авторизации)
        private void BtnShowDebts_Click(object sender, RoutedEventArgs e)
        {
            DebtsFilterWindow debtsWindow = new DebtsFilterWindow();
            debtsWindow.ShowDialog();
        }
    }
}
