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
using System.Windows.Shapes;
using Tyreshop.DbAccess;
using NLog;
using System.Threading;

namespace Tyreshop
{
    /// <summary>
    /// Логика взаимодействия для UserManageWindow.xaml
    /// </summary>
    public partial class UserManageWindow : Window
    {
        private Logger log;
        private MainWindow _mainWnd;
        public UserManageWindow(MainWindow wnd)
        {
            InitializeComponent();
            _mainWnd = wnd;
            log = LogManager.GetCurrentClassLogger();
            FillLoadControls();
        }

        private void FillLoadControls() {
            try {
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities()) {
                    var users = db.users.ToList();
                    Users.ItemsSource = users;
                    SelectedUser.ItemsSource = users;
                    SelectedUser.SelectedValuePath = "UserId";
                    SelectedUser.DisplayMemberPath = "UserName";
                    var roles = new List<string>() { "admin", "manager", "cashier", "buyer" };
                    Role.ItemsSource = roles;
                    RoleEdit.ItemsSource = roles;
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void AddUser_Click(object sender, RoutedEventArgs e)
        {
            bool Flag = true;
            string userName = "", login = "", pass = "", role = "";
            try {
                if (UserName.Text != string.Empty)
                    userName = UserName.Text;
                else
                    Flag = false;
                if (Login.Text != string.Empty)
                    login = Login.Text;
                else
                    Flag = false;
                if (Password.Text != string.Empty && PasswordRep.Text != string.Empty && Password.Text == PasswordRep.Text)
                    pass = Password.Text;
                else
                    Flag = false;
                if (Role.SelectedValue != null)
                    role = (string)Role.SelectedValue;
                else
                    Flag = false;
                if (Flag) {
                    try {
                        using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities()) {
                            if (!db.users.Any(a => a.Login == login && a.Password == pass))
                            {
                                var newUser = new user()
                                {
                                    UserName = userName,
                                    Login = login,
                                    Password = pass,
                                    Role = role
                                };
                                db.users.Add(newUser);
                                db.SaveChanges();
                                MessageBox.Show("Пользователь успешно добавлен", "Информация", MessageBoxButton.OK);
                                Thread.Sleep(1000);
                                FillLoadControls();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        log.Error(ex.Message + " \n" + ex.StackTrace);
                        MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
                    }
                    UserName.Text = "";
                    Login.Text = "";
                    Password.Text = "";
                    PasswordRep.Text = "";
                    Role.Text = "";
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void SelectedUser_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var cmb = sender as ComboBox;
            try {
                int userId = (int)cmb.SelectedValue;
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities()) {
                    var user = db.users.Find(userId);
                    LoginEdit.Text = user.Login;
                    PasswordNow.Text = user.Password;
                    PasswordNew.Text = user.Password;
                    RoleEdit.SelectedValue = user.Role;
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void DelUser_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedUser.SelectedValue != null)
            {
                int userId = (int)SelectedUser.SelectedValue;
                var res = MessageBox.Show("Вы действительно хотите полностью удалить этого пользователя? Действие необратимо.", "Информация", MessageBoxButton.OKCancel);
                if (res == MessageBoxResult.OK)
                {
                    try
                    {
                        using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                        {
                            var user = db.users.Single(s => s.UserId == userId);
                            db.users.Remove(user);
                            db.SaveChanges();

                        }
                    }
                    catch (Exception ex)
                    {
                        log.Error(ex.Message + " \n" + ex.StackTrace);
                        MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
                    }
                    SelectedUser.SelectionChanged -= SelectedUser_SelectionChanged;
                    SelectedUser.Text = "";
                    SelectedUser.SelectedValue = -1;
                    SelectedUser.SelectionChanged += SelectedUser_SelectionChanged;
                    LoginEdit.Text = "";
                    PasswordNow.Text = "";
                    PasswordNew.Text = "";
                    RoleEdit.Text = "";
                }
                MessageBox.Show("Пользователь успешно удален", "Информация", MessageBoxButton.OK);
                Thread.Sleep(1000);
                FillLoadControls();
            }
        }

        private void EditUser_Click(object sender, RoutedEventArgs e)
        {
            bool Flag = false;
            if (SelectedUser.SelectedValue != null)
            {
                int userId = (int)SelectedUser.SelectedValue;
                try
                {
                    using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                    {
                        var user = db.users.Single(s => s.UserId == userId);

                        if (LoginEdit.Text != user.Login) {
                            user.Login = LoginEdit.Text;
                            db.Entry(user).Property(p => p.Login).IsModified = true;
                            Flag = true;
                        }

                        if (PasswordNow.Text == user.Password && PasswordNew.Text != user.Password && PasswordNew.Text != string.Empty)
                        {
                            user.Password = PasswordNew.Text;
                            db.Entry(user).Property(p => p.Password).IsModified = true;
                            Flag = true;
                        }
                        else {
                            MessageBox.Show("Вы выбрали небезопасный пароль, или не ввели его", "Информация", MessageBoxButton.OK);
                        }
                        if (RoleEdit.SelectedValue != null)
                        {
                            user.Role = (string)RoleEdit.SelectedValue;
                            db.Entry(user).Property(p => p.Role).IsModified = true;
                            Flag = true;
                        }
                        if (Flag)
                        {
                            bool flag = true;
                            try
                            {
                                db.SaveChanges();
                            }
                            catch (Exception ex)
                            {
                                log.Error(ex.Message + " \n" + ex.StackTrace);
                                flag = false;
                                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
                            }
                            if (flag)
                            {
                                SelectedUser.SelectionChanged -= SelectedUser_SelectionChanged;
                                SelectedUser.Text = "";
                                SelectedUser.SelectedValue = -1;
                                SelectedUser.SelectionChanged += SelectedUser_SelectionChanged;
                                LoginEdit.Text = "";
                                PasswordNow.Text = "";
                                PasswordNew.Text = "";
                                RoleEdit.Text = "";
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    log.Error(ex.Message + " \n" + ex.StackTrace);
                    MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
                }
                MessageBox.Show("Пользователь успешно отредактирован", "Информация", MessageBoxButton.OK);
                Thread.Sleep(1000);
                FillLoadControls();
            }
        }
    }
}
