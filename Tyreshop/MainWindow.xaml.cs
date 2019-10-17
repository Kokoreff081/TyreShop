using System;
using System.Collections.Generic;
using System.Configuration;
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
using MySql.Data.MySqlClient;
using System.Security.Cryptography;

namespace Tyreshop
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public user User;
        public MainWindow()
        {
            InitializeComponent();
            this.Height = (System.Windows.SystemParameters.PrimaryScreenHeight-50);
            this.Width = (System.Windows.SystemParameters.PrimaryScreenWidth);
            OpenPages(pages.login);

        }
        public enum pages {
            login,
            register,
            main
        }
        private void OpenPages(pages page) {
            if (page == pages.login)
                mainFrame.Navigate(new login(this));
            else
                mainFrame.Navigate(pages.main);
        }
    }
}
