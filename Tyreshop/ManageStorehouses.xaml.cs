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
using Tyreshop.Utils;

namespace Tyreshop
{
    /// <summary>
    /// Логика взаимодействия для ManageStorehouses.xaml
    /// </summary>
    public partial class ManageStorehouses : Window
    {
        public ManageStorehouses()
        {
            InitializeComponent();
            LoadSelectStore();
        }

        private void LoadSelectStore() {
            using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities()) {
                var stores = db.storehouses.ToList();
                SelectStorehouse.ItemsSource = stores;
                SelectStorehouse.SelectedValuePath = "StorehouseId";
                SelectStorehouse.DisplayMemberPath = "StorehouseName";
            }
        }

        private void AddStorehouse_Click(object sender, RoutedEventArgs e)
        {
            string name = StorehouseName.Text;
            string address = Address.Text;
            if (name != string.Empty && address != string.Empty) {
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities()) {
                    storehouse sh = new storehouse()
                    {
                        StorehouseName = name,
                        Address = address
                    };
                    db.storehouses.Add(sh);
                    db.SaveChanges();
                    MessageBox.Show("Склад успешно добавлен!", "Информация", MessageBoxButton.OK);
                }
            }
            else {
                MessageBox.Show("Не заполнено название или адрес нового склада!", "Информация", MessageBoxButton.OK);
            }
            LoadSelectStore();
        }

        private void SelectStorehouse_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var cmb = sender as ComboBox;
            int id = (int)cmb.SelectedValue;
            using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities()) {
                var list = new List<StoreQuant>();
                var prodQuant = db.productquantities.Where(w => w.StorehouseId == id).ToList();
                var dbProds = db.products.ToList();
                var prodsInStore = dbProds.Where(w=>prodQuant.Any(a=>a.ProductId==w.ProductId))
                    .GroupJoin(prodQuant, dp => dp.ProductId, pq => pq.ProductId, (dp, pq) => new StoreQuant() {
                        PSeason = dp.Season,
                        PQuant = (int)pq.Sum(x => x.Quantity),
                    })
                    .GroupBy(g=>g.PSeason)
                    .ToDictionary(k=>k.Key);
                foreach (var item in prodsInStore) {
                    StoreQuant sq = new StoreQuant();
                    sq.PSeason = item.Key;
                    sq.TotalQuant = item.Value.Sum(s => s.PQuant);
                    list.Add(sq);
                    int point2 = 0;
                }
                int point = 0;
                //var prods = 
                QuantityAtStore.ItemsSource = list;
            }
        }
    }
}
