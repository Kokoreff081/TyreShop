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

namespace Tyreshop
{
    /// <summary>
    /// Логика взаимодействия для AddService.xaml
    /// </summary>
    public partial class AddService : Window
    {
        public AddService()
        {
            InitializeComponent();
        }

        private void AddServiceBtn_Click(object sender, RoutedEventArgs e)
        {
            string name = ServiceName.Text;
            if (name != string.Empty)
            {
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    service newMan = new service()
                    {
                        ServiceName = name
                    };
                    db.services.Add(newMan);
                    db.SaveChanges();
                    MessageBox.Show("Услуга успешно добавлена!", "Информация", MessageBoxButton.OK);
                }
            }
            else
            {
                MessageBox.Show("Введите наименование услуги!", "Информация", MessageBoxButton.OK);
            }
        }
    }
}
