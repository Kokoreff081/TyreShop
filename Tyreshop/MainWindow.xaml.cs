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
using System.Reflection;
using System.Deployment.Application;

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
            Assembly assem = Assembly.GetEntryAssembly();
            AssemblyName assemName = assem.GetName();
            Version ver = assemName.Version;
            string version = "Система склад ( V. " + getRunningVersion() + " )";
            MainWnd.Title = version;
            this.Height = (System.Windows.SystemParameters.PrimaryScreenHeight-50);
            this.Width = (System.Windows.SystemParameters.PrimaryScreenWidth);
            OpenPages(pages.login);

        }

        private Version getRunningVersion()
        {
            try
            {
                return Assembly.GetExecutingAssembly().GetName().Version;
            }
            catch (Exception)
            {
                return ApplicationDeployment.CurrentDeployment.CurrentVersion;
            }
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
