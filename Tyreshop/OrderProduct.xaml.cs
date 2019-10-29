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
using NLog;

namespace Tyreshop
{
    /// <summary>
    /// Логика взаимодействия для OrderProduct.xaml
    /// </summary>
    public partial class OrderProduct : Window
    {
        private Logger log;
        private List<PComboBox> list;
        private List<BdProducts> _MainList;
        private user _user;
        public OrderProduct(user user, List<PComboBox> lpc, List<BdProducts> bdp)
        {
            InitializeComponent();
            log = LogManager.GetCurrentClassLogger();
            _user = user;
            list = lpc;
            _MainList = bdp;
            LoadAddOrderForm();
            LoadOrderList();
        }

        private void LoadAddOrderForm() {
                Products.ItemsSource = list;
                Products.SelectedValuePath = "ProductId";
                Products.DisplayMemberPath = "ProductName";
        }

        private void LoadOrderList() {
            List<OrderList> lst = new List<OrderList>();
            try
            {
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    var orders = db.Orders.ToList();
                    foreach (var order in orders)
                    {
                        var user = db.users.Single(s => s.UserId == order.MengerId);
                        OrderList item = new OrderList()
                        {
                            OrderId = order.OrderID,
                            OrderCustomer = order.OrderCustomer,
                            OrderNumber = order.OrderNumber,
                            PhoneCustomer = order.PhoneCustomer,
                            Quantity = order.ProductQuant,
                            ProdName = list.Where(w => w.ProductId == order.ProductId).Select(s => s.ProductName).First(),
                            Storehouse = db.storehouses.Where(w => w.StorehouseId == order.StorehouseId).Select(s => s.StorehouseName).First(),
                            UserName = user.UserName,
                            UserRole = user.Role
                        };
                        lst.Add(item);
                    }
                }
                Orders.ItemsSource = lst;
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void Products_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var Cmb = sender as ComboBox;
            try
            {
                int prodId = (int)Cmb.SelectedValue;
                decimal prodPrice = 0;
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    prodPrice = (decimal)_MainList.Where(w => w.ProductId == prodId).Select(s => s.Price).FirstOrDefault();
                    var storehouses = db.productquantities.Where(w => w.ProductId == prodId && w.Quantity>0).Join(db.storehouses, pq => pq.StorehouseId, s => s.StorehouseId, (pq, s) => new { StoreHouseId = pq.StorehouseId, StoreHouseName = s.StorehouseName }).ToList();
                    Storehouse.ItemsSource = storehouses;
                    Storehouse.SelectedValuePath = "StoreHouseId";
                    Storehouse.DisplayMemberPath = "StoreHouseName";
                    Storehouse.IsEnabled = true;
                }
                ProductPrice.Text = prodPrice.ToString();
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void Products_KeyUp(object sender, KeyEventArgs e)
        {
            var Cmb = sender as ComboBox;
            try
            {
                CollectionView itemsViewOriginal = (CollectionView)CollectionViewSource.GetDefaultView(Cmb.ItemsSource);
                itemsViewOriginal.Filter = ((o) =>
                {
                    if (String.IsNullOrEmpty(Cmb.Text)) return true;
                    else
                    {
                        var obj = o as PComboBox;
                        if ((obj.ProductName).Contains(Cmb.Text))
                        {
                            Cmb.IsDropDownOpen = true;
                            if (Key.Enter == e.Key)
                            {
                                int prodId = (int)Cmb.SelectedValue;
                                decimal prodPrice = 0;
                                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                                {
                                    prodPrice = (decimal)_MainList.Where(w => w.ProductId == prodId).Select(s => s.Price).FirstOrDefault();
                                    var storehouses = db.productquantities.Where(w => w.ProductId == prodId).Join(db.storehouses, pq => pq.StorehouseId, s => s.StorehouseId, (pq, s) => new { StoreHouseId = pq.StorehouseId, StoreHouseName = s.StorehouseName }).ToList();
                                    Storehouse.ItemsSource = storehouses;
                                    Storehouse.SelectedValuePath = "StoreHouseId";
                                    Storehouse.DisplayMemberPath = "StoreHouseName";
                                    Storehouse.IsEnabled = true;
                                }
                                ProductPrice.Text = prodPrice.ToString();
                            }
                            return true;
                        }
                        else return false;
                    }
                });
                itemsViewOriginal.Refresh();
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void AddOrder_Click(object sender, RoutedEventArgs e)
        {
            bool flag = false;
            try
            {
                if (Products.SelectedValue != null)
                {
                    flag = true;
                    int quant = 0, storeId = 0;
                    string customerName = Customer.Text;
                    string customerPhone = Phone.Text;
                    int prodId = (int)Products.SelectedValue;
                    if (int.TryParse(OrderQuant.Text, out var tmp))
                    {
                        quant = int.Parse(OrderQuant.Text);
                    }
                    else
                        flag = false;
                    if (Storehouse.SelectedValue != null)
                    {
                        storeId = (int)Storehouse.SelectedValue;
                    }
                    else
                        flag = false;
                    if (flag == true)
                    {
                        using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                        {
                            var prodQuant = db.productquantities.Single(s => s.ProductId == prodId && s.StorehouseId == storeId);
                            var diff = prodQuant.Quantity - quant;
                            if (diff >= 0)
                            {
                                var sNum = db.Orders.ToList();
                                long number = 0;
                                if (sNum.Count == 0)
                                    number = 1;
                                else
                                    number = (long)sNum[sNum.Count - 1].OrderNumber + 1;
                                Order newOrd = new Order()
                                {
                                    OrderNumber = number,
                                    OrderCustomer = customerName,
                                    PhoneCustomer = customerPhone,
                                    ProductId = prodId,
                                    ProductQuant = quant,
                                    StorehouseId = storeId,
                                    MengerId = _user.UserId
                                };
                                db.Orders.Add(newOrd);
                                if (prodQuant.InOrder != null)
                                    prodQuant.InOrder += quant;
                                else
                                    prodQuant.InOrder = quant;
                                db.Entry(prodQuant).Property(p => p.InOrder).IsModified = true;
                                db.SaveChanges();
                            }
                            else
                                MessageBox.Show("На выбранном складе недостаточное количество товара для осуществления брони!", "Информация", MessageBoxButton.OK);
                        }
                        LoadOrderList();
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }

        }

        private void DelOrder_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            try
            {
                if (long.TryParse(btn.Uid, out var tmp))
                {
                    long ordId = long.Parse(btn.Uid);
                    long ordNum = (long)btn.Tag;
                    using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                    {
                        var order = db.Orders.Single(s => s.OrderID == ordId && s.OrderNumber == ordNum);
                        int quant = order.ProductQuant;
                        db.Orders.Remove(order);
                        var prodQuant = db.productquantities.Single(s => s.ProductId == order.ProductId && s.StorehouseId == order.StorehouseId);
                        prodQuant.InOrder -= quant;
                        db.Entry(prodQuant).Property(p => p.InOrder).IsModified = true;
                        db.SaveChanges();
                        MessageBox.Show("Бронь успешно удалена!", "Информация", MessageBoxButton.OK);
                        LoadOrderList();
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }
    }
}
