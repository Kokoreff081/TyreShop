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
using Tyreshop.DbAccess;
using NLog;

namespace Tyreshop
{
    /// <summary>
    /// Логика взаимодействия для login.xaml
    /// </summary>
    public partial class login : Page
    {
        public MainWindow _mainWnd;
        private Logger log;
        public login(MainWindow mainWnd)
        {
            InitializeComponent();
            _mainWnd = mainWnd;
            log = LogManager.GetCurrentClassLogger(); 
        }

        private void LoginBtn_Click(object sender, RoutedEventArgs e)
        {
            if (Login.Text.Length > 0)
            {
                if (Password.Password.Length > 0)
                {
                    var login = Login.Text;
                    var pass = Password.Password;
                    try
                    {
                        using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                        {
                            var user = db.users.ToList().Single(w => w.Login == login && w.Password == pass);
                            if (user != null)
                            {
                                _mainWnd.User = user;
                                _mainWnd.mainFrame.Navigate(new main(_mainWnd, user));

                            }
                        }
                    }
                    catch (Exception ex) {
                        log.Error(ex.Message + " \n" + ex.StackTrace);
                        MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);

                    }
                   
                }
            }
        }

        private void RegisterBtn_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
