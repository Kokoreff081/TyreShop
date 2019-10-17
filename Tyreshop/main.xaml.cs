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
using Tyreshop.Utils;
using MoreLinq;
using Microsoft.Win32;
using System.Windows.Threading;
using NLog;

namespace Tyreshop
{
    /// <summary>
    /// Логика взаимодействия для main.xaml
    /// </summary>
    public partial class main : Page
    {
        private MainWindow _mainWnd;
        private user _user;
        internal List<BdProducts> MainList;
        private List<BdProducts> tmpList;
        private List<BdProducts> filtered;
        private List<ComboBox> tbFilters;
        private List<BdProducts> tmp;
        internal List<PComboBox> listPCombo;
        private Logger log;
        private List<string> heights = new List<string>();
        private List<string> widths = new List<string>();
        private List<string> radiuses = new List<string>();
        public main(MainWindow mWnd, user user)
        {
            InitializeComponent();
            log = LogManager.GetCurrentClassLogger();
            _mainWnd = mWnd;
            _user = user;
            if (_user.Role == "admin")
                Dashboard.Visibility = Visibility.Visible;
            if (_user.Role == "cashier" || _user.Role=="admin")
            {
                SaleReceipt.Visibility = Visibility.Visible;
                DailyReport.Visibility = Visibility.Visible;
            }
            MainList = GetItemSource();
            FillLoadControls();
            tbFilters = new List<ComboBox>();
            tbFilters.Add(Width);
            tbFilters.Add(Height);
            tbFilters.Add(Manufacturer);
            tbFilters.Add(Model);
            tbFilters.Add(IsCol);
            tbFilters.Add(InCol);
            tbFilters.Add(RFT);
            tbFilters.Add(Season);
            tbFilters.Add(Gruz);
            tbFilters.Add(Radius);
            tmp = new List<BdProducts>();
            try
            {
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    listPCombo = db.products.Select(s => new PComboBox()
                    {
                        ProductId = s.ProductId,
                        ProductName = @"" + db.manufacturers.Where(w => w.ManufacturerId == s.ManufacturerId).Select(sel => sel.ManufacturerName.Trim()).FirstOrDefault() + " " +
                            db.models.Where(w => w.ModelId == s.ModelId && w.ManufacturerId == s.ManufacturerId).Select(sel => sel.ModelName.Trim()).FirstOrDefault() + " " + s.Width + "/" + s.Height + "/"
                            + s.Radius + "/" + s.RFT + "/" + s.IsCol + "/" + s.InCol
                    })
                    .OrderBy(o => o.ProductName)
                    .ToList();
                }
            }
            catch (Exception e) {
                log.Error(e.Message + " \n" + e.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
            
        }

        private void FillLoadControls() {
            List<TreeViewContent> mainList = new List<TreeViewContent>();
            List<GroupOfProducts> radiusList = new List<GroupOfProducts>();
            List<BdProducts> products = new List<BdProducts>();
            try
            {
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    widths = db.products.Select(s => s.Width.ToString()).Distinct().ToList();
                    Width.ItemsSource = widths;
                    heights = db.products.Select(s => s.Height.ToString()).Distinct().ToList();
                    Height.ItemsSource = heights;
                    radiuses = db.products.Select(s => s.Radius).Distinct().ToList();
                    Radius.ItemsSource = radiuses;
                    var manufs = db.products.Join(db.manufacturers, p => p.ManufacturerId, m => m.ManufacturerId, (p, m) => new { Name = m.ManufacturerName.Trim() }).Select(s => s.Name).Distinct().OrderBy(o => o).ToList();
                    Manufacturer.ItemsSource = manufs;
                    var models = db.products.Join(db.models, p => p.ModelId, m => m.ModelId, (p, m) => new { Name = m.ModelName }).Select(s => s.Name).Distinct().OrderBy(o => o).ToList();
                    Model.ItemsSource = models;
                    var isCols = db.products.Select(s => s.IsCol).Distinct().ToList();
                    IsCol.ItemsSource = isCols;
                    var inCols = db.products.Select(s => s.InCol).Distinct().ToList();
                    InCol.ItemsSource = inCols;
                    var rfts = db.products.Select(s => s.RFT).Distinct().ToList();
                    RFT.ItemsSource = rfts;
                    var gruzes = db.products.Select(s => s.Gruz).Distinct().ToList();
                    Gruz.ItemsSource = gruzes;
                    var seasons = db.products.Select(s => s.Season).Distinct().ToList();
                    Season.ItemsSource = seasons;
                    var spikes = db.products.Select(s => s.Spikes).Distinct().ToList();
                    Spikes.ItemsSource = spikes;
                    var groups = db.products.Select(s => s.Radius).Distinct().ToList();
                    foreach (var g in groups)
                    {
                        GroupOfProducts gop = new GroupOfProducts()
                        {
                            Radius = g,
                            BdProducts = MainList.Where(w => w.Radius == g).ToList()
                        };
                        radiusList.Add(gop);
                    }
                    var categories = db.categories.ToList();
                    foreach (var c in categories)
                    {
                        TreeViewContent tvc = new TreeViewContent()
                        {
                            CategoryId = c.CategoryId,
                            CategoryName = c.CategoryName,
                            Products = radiusList.Where(w => w.BdProducts.Any(a => a.CategoryId == c.CategoryId)).ToList()
                        };
                        mainList.Add(tvc);
                    }
                }
                ProductTV.ItemsSource = mainList;
            }
            catch (Exception e) {
                log.Error(e.Message + " \n" + e.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
            
        }

        private void AddProduct_Click(object sender, RoutedEventArgs e)
        {
            AddProduct wnd = new AddProduct(_mainWnd, MainList, listPCombo);
            wnd.Owner = _mainWnd;
            wnd.Show();
        }

        private void CapitalizeProduct_Click(object sender, RoutedEventArgs e)
        {
            CapitalizeProduct wnd = new CapitalizeProduct(_mainWnd, listPCombo, MainList);
            wnd.Owner = _mainWnd;
            wnd.Show();
        }

        private void AddStorehouse_Click(object sender, RoutedEventArgs e)
        {
            ManageStorehouses wnd = new ManageStorehouses();
            wnd.Owner = _mainWnd;
            wnd.Show();
        }

        private void AddUser_Click(object sender, RoutedEventArgs e)
        {
            UserManageWindow wnd = new UserManageWindow(_mainWnd);
            wnd.Owner = _mainWnd;
            wnd.Show();
        }

        private void AddCategory_Click(object sender, RoutedEventArgs e)
        {

        }

        private void ProductsExpostXlsx_Click(object sender, RoutedEventArgs e)
        {
            
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "Excel file (*.xlsx)|*.xlsx";
                if (sfd.ShowDialog() == true)
                {
                    try
                    {
                        XlsxExport.ExportToSite(MainList, sfd);
                    }
                    catch (Exception ex) {
                        string msg = "Выбранный файл недоступен: " + ex.Message + "/r/nПожалуйста, повторите сохранение, задав другое имя файла.";
                        MessageBoxResult res = MessageBox.Show(msg, "Информация", MessageBoxButton.OK);
                    }
                }
            
        }

        private void SaleReceipt_Click(object sender, RoutedEventArgs e)
        {
            SaleFixing wnd = new SaleFixing(_mainWnd, listPCombo,MainList);
            wnd.Owner = _mainWnd;
            wnd.Show();
        }

        private void ProductTV_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            TreeView item = sender as TreeView;
            TreeViewItem selectedItem = ContainerFromItem(ProductTV.ItemContainerGenerator, item.SelectedItem);
            string type = selectedItem.Uid;
            if (type == "Radius")
            {
                string rad = selectedItem.Tag.ToString();
                tmpList = MainList.Where(w=>w.Radius==rad).ToList();
                Details.ItemsSource = tmpList;
                if (_user.Role != "admin")
                    Details.Columns[14].Visibility = Visibility.Collapsed;
                Radius.Visibility = Visibility.Collapsed;
                RadiusLbl.Visibility = Visibility.Collapsed;
            }
            else {
                tmpList = MainList;
                Details.ItemsSource = MainList;
                if (_user.Role != "admin")
                    Details.Columns[14].Visibility = Visibility.Collapsed;
                Radius.Visibility = Visibility.Visible;
                RadiusLbl.Visibility = Visibility.Visible;
            }
        }

        private List<BdProducts> GetItemSource() {
            List<BdProducts> prods = new List<BdProducts>();
            try
            {
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    prods = db.products.GroupJoin(db.productquantities, p => p.ProductId, pq => pq.ProductId, (p, pq) => new BdProducts()
                    {
                        ProductId = p.ProductId,
                        ModelId = p.ModelId,
                        Articul = p.Articul,
                        Gruz = p.Gruz,
                        Height = p.Height,
                        InCol = p.InCol,
                        IsCol = p.IsCol,
                        Manufacturer = db.manufacturers.Where(s => s.ManufacturerId == p.ManufacturerId).Select(s => s.ManufacturerName).FirstOrDefault(),
                        Model = db.models.Where(w => w.ModelId == p.ModelId).Select(s => s.ModelName).FirstOrDefault(),
                        Radius = p.Radius,
                        RFT = p.RFT,
                        Season = p.Season,
                        Spikes = p.Spikes,
                        Width = p.Width,
                        Price = (decimal)p.Price,
                        OptPrice = p.OptPrice,
                        PurPrice = p.PurchasePrice,
                        CategoryId = p.CategoryId,
                        Country = p.Country,
                        TotalQuantity = pq.Sum(x => x.Quantity),
                        InOrder = pq.Sum(s => s.InOrder),
                        Storehouse = db.productquantities.Where(w => w.ProductId == p.ProductId).Join(db.storehouses, pq1 => pq1.StorehouseId, s => s.StorehouseId, (pq1, s) => new ProductQuantity()
                        {
                            Id = pq1.Id,
                            ProductId = pq1.ProductId,
                            Quantity = (int)pq1.Quantity,
                            InOrder = pq1.InOrder,
                            Address = s.Address,
                            StorehouseName = s.StorehouseName
                        }).ToList()//pq.StorehouseId
                    })
                        .OrderBy(o => o.Manufacturer.Trim())
                        .ThenBy(t => t.Model.Trim())
                        .ThenBy(t => t.Model.Trim())
                        .ThenBy(t => t.Width)
                        .ThenBy(t => t.Height)
                        .ThenBy(t => t.Radius)
                        .ToList();
                }
            }
            catch (Exception e) { log.Error(e.Message + " \n" + e.StackTrace); MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK); }
            return prods;
        }

        private List<BdProducts> FilterMainList(List<BdProducts> list, string rad = "", string width = "", string height = "", string manId = "", string modelId = "", string IsCol = "", string InCol = "", string Gruz = "", string RFT = "", string Spikes = "", string season = "") {
            var prods = list;
            if (!string.IsNullOrEmpty(rad))
                prods = prods.Where(w => w.Radius == rad).ToList();
            if (!string.IsNullOrEmpty(width))
                prods = prods.Where(w => w.Width.ToString() == width).ToList();
            if (!string.IsNullOrEmpty(height))
                prods = prods.Where(w => w.Height.ToString() == height).ToList();
            if (!string.IsNullOrEmpty(manId))
                prods = prods.Where(w => w.Manufacturer == manId).ToList();
            if (!string.IsNullOrEmpty(modelId))
                prods = prods.Where(w => w.Model == modelId).ToList();
            if (!string.IsNullOrEmpty(IsCol))
                prods = prods.Where(w => w.IsCol == IsCol).ToList();
            if (!string.IsNullOrEmpty(InCol))
                prods = prods.Where(w => w.InCol == InCol).ToList();
            if (!string.IsNullOrEmpty(Gruz))
                prods = prods.Where(w => w.Gruz == Gruz).ToList();
            if (!string.IsNullOrEmpty(RFT))
                prods = prods.Where(w => w.RFT == RFT).ToList();
            if (!string.IsNullOrEmpty(Spikes))
                prods = prods.Where(w => w.Spikes == Spikes).ToList();
            if (!string.IsNullOrEmpty(season))
                prods = prods.Where(w => w.Season == season).ToList();

            return prods;
        }

        private static TreeViewItem ContainerFromItem(ItemContainerGenerator containerGenerator, object item)
        {
            TreeViewItem container = (TreeViewItem)containerGenerator.ContainerFromItem(item);
            if (container != null)
                return container;

            foreach (object childItem in containerGenerator.Items)
            {
                TreeViewItem parent = containerGenerator.ContainerFromItem(childItem) as TreeViewItem;
                if (parent == null)
                    continue;

                container = parent.ItemContainerGenerator.ContainerFromItem(item) as TreeViewItem;
                if (container != null)
                    return container;

                container = ContainerFromItem(parent.ItemContainerGenerator, item);
                if (container != null)
                    return container;
            }
            return null;
        }

        private void AddService_Click(object sender, RoutedEventArgs e)
        {
            AddService wnd = new AddService();
            wnd.Owner = _mainWnd;
            wnd.Show();
        }

        private void DailyReport_Click(object sender, RoutedEventArgs e)
        {
            DailyReportWindow wnd = new DailyReportWindow(listPCombo);
            wnd.Owner = _mainWnd;
            wnd.Show();
        }

        private void Orders_Click(object sender, RoutedEventArgs e)
        {
            OrderProduct wnd = new OrderProduct(_user, listPCombo, MainList);
            wnd.Owner = _mainWnd;
            wnd.Show();
        }

        private void Width_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var tb = sender as ComboBox;
            var prods = MainList;//Details.ItemsSource as List<BdProducts>;
            if (tb.SelectedValue != null)
            {
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    var heights1 = db.products.Where(w => w.Width.ToString() == (string)tb.SelectedValue).Select(s => s.Height.ToString()).Distinct().ToList();
                    Height.ItemsSource = heights1;
                    var radiuses1 = db.products.Where(w => w.Width.ToString() == (string)tb.SelectedValue).Select(s => s.Radius).Distinct().ToList();
                    Radius.ItemsSource = radiuses1;
                }
                if (prods.Count > 0)
                {
                    if (Height.SelectedValue != null || Season.SelectedValue != null || Manufacturer.SelectedValue != null || RFT.SelectedValue != null || IsCol.SelectedValue != null ||
                        InCol.SelectedValue != null || Model.SelectedValue != null || Gruz.SelectedValue != null || Radius.SelectedValue != null || Spikes.SelectedValue != null)
                        filtered = FilterMainList(MainList, (string)Radius.SelectedValue, (string)tb.SelectedValue, (string)Height.SelectedValue, (string)Manufacturer.SelectedValue, (string)Model.SelectedValue, (string)IsCol.SelectedValue, (string)InCol.SelectedValue, (string)Gruz.SelectedValue, (string)RFT.SelectedValue, (string)Spikes.SelectedValue, (string)Season.SelectedValue);
                    else
                        filtered = FilterMainList(MainList, null, (string)tb.SelectedValue, null, null, null, null, null, null, null, null, null);
                    Details.ItemsSource = filtered;
                }
                
            }
            else
            {
                if (Height.SelectedValue != null || Season.SelectedValue != null || Manufacturer.SelectedValue != null || RFT.SelectedValue != null || IsCol.SelectedValue != null ||
                        InCol.SelectedValue != null || Model.SelectedValue != null || Gruz.SelectedValue != null || Radius.SelectedValue != null || Spikes.SelectedValue != null)
                    filtered = FilterMainList(MainList, (string)Radius.SelectedValue, (string)tb.SelectedValue, (string)Height.SelectedValue, (string)Manufacturer.SelectedValue, (string)Model.SelectedValue, (string)IsCol.SelectedValue, (string)InCol.SelectedValue, (string)Gruz.SelectedValue, (string)RFT.SelectedValue, (string)Spikes.SelectedValue, (string)Season.SelectedValue);
                else
                    Details.ItemsSource = tmpList;
            }
        }

        private void Height_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var tb = sender as ComboBox;
            var prods = MainList;//Details.ItemsSource as List<BdProducts>;
            if (tb.SelectedValue != null)
            {
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    //var widths = db.products.Where(w => w.Height.ToString() == (string)tb.SelectedValue).Select(s => s.Height.ToString()).Distinct().ToList();
                    //Width.ItemsSource = widths;
                    var radiuses1 = db.products.Where(w => w.Height.ToString() == (string)tb.SelectedValue).Select(s => s.Radius).Distinct().ToList();
                    Radius.ItemsSource = radiuses1;
                }
                if (prods.Count > 0)
                {
                    if (Width.SelectedValue != null || Season.SelectedValue != null || Manufacturer.SelectedValue != null || RFT.SelectedValue != null || IsCol.SelectedValue != null ||
                        InCol.SelectedValue != null || Model.SelectedValue != null || Gruz.SelectedValue != null || Radius.SelectedValue != null || Spikes.SelectedValue != null)
                        filtered = FilterMainList(MainList, (string)Radius.SelectedValue, (string)Width.SelectedValue, (string)tb.SelectedValue, (string)Manufacturer.SelectedValue, (string)Model.SelectedValue, (string)IsCol.SelectedValue, (string)InCol.SelectedValue, (string)Gruz.SelectedValue, (string)RFT.SelectedValue, (string)Spikes.SelectedValue, (string)Season.SelectedValue);
                    else
                        filtered = FilterMainList(MainList, null, null, (string)tb.SelectedValue, null, null, null, null, null, null, null, null);
                    Details.ItemsSource = filtered;
                }
            }
            else
            {
                if (Width.SelectedValue != null || Season.SelectedValue != null || Manufacturer.SelectedValue != null || RFT.SelectedValue != null || IsCol.SelectedValue != null ||
                        InCol.SelectedValue != null || Model.SelectedValue != null || Gruz.SelectedValue != null || Radius.SelectedValue != null || Spikes.SelectedValue != null)
                    filtered = FilterMainList(MainList, (string)Radius.SelectedValue, (string)Width.SelectedValue, (string)tb.SelectedValue, (string)Manufacturer.SelectedValue, (string)Model.SelectedValue, (string)IsCol.SelectedValue, (string)InCol.SelectedValue, (string)Gruz.SelectedValue, (string)RFT.SelectedValue, (string)Spikes.SelectedValue, (string)Season.SelectedValue);
                else
                    Details.ItemsSource = tmpList;
            }
        }

        private void Radius_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var tb = sender as ComboBox;
            var prods = MainList;//Details.ItemsSource as List<BdProducts>;
            if (tb.SelectedValue != null)
            {
                //using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                //{
                //    var heights = db.products.Where(w => w.Radius == (string)tb.SelectedValue).Select(s => s.Height.ToString()).Distinct().ToList();
                //    Height.ItemsSource = heights;
                //    var widths = db.products.Where(w => w.Radius == (string)tb.SelectedValue).Select(s => s.Width).Distinct().ToList();
                //    Width.ItemsSource = widths;
                //}
                if (prods.Count > 0)
                {
                    if (Width.SelectedValue != null || Season.SelectedValue != null || Manufacturer.SelectedValue != null || RFT.SelectedValue != null || IsCol.SelectedValue != null ||
                        InCol.SelectedValue != null || Model.SelectedValue != null || Gruz.SelectedValue != null || Height.SelectedValue != null || Spikes.SelectedValue != null)
                        filtered = FilterMainList(MainList, (string)tb.SelectedValue, (string)Width.SelectedValue, (string)Height.SelectedValue, (string)Manufacturer.SelectedValue, (string)Model.SelectedValue, (string)IsCol.SelectedValue, (string)InCol.SelectedValue, (string)Gruz.SelectedValue, (string)RFT.SelectedValue, (string)Spikes.SelectedValue, (string)Season.SelectedValue);
                    else
                        filtered = FilterMainList(MainList, (string)tb.SelectedValue, null, null, null, null, null, null, null, null, null, null);
                    Details.ItemsSource = filtered;
                }
            }
            else
            {
                if (Width.SelectedValue != null || Season.SelectedValue != null || Manufacturer.SelectedValue != null || RFT.SelectedValue != null || IsCol.SelectedValue != null ||
                       InCol.SelectedValue != null || Model.SelectedValue != null || Gruz.SelectedValue != null || Height.SelectedValue != null || Spikes.SelectedValue != null)
                    filtered = FilterMainList(MainList, (string)tb.SelectedValue, (string)Width.SelectedValue, (string)Height.SelectedValue, (string)Manufacturer.SelectedValue, (string)Model.SelectedValue, (string)IsCol.SelectedValue, (string)InCol.SelectedValue, (string)Gruz.SelectedValue, (string)RFT.SelectedValue, (string)Spikes.SelectedValue, (string)Season.SelectedValue);
                else
                    Details.ItemsSource = tmpList;
            }
        }

        private void Season_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var tb = sender as ComboBox;
            var prods = MainList;//Details.ItemsSource as List<BdProducts>;
            if (tb.SelectedValue != null)
            {
                if (prods.Count > 0)
                {
                    if (Height.SelectedValue != null || Gruz.SelectedValue != null || Manufacturer.SelectedValue != null || RFT.SelectedValue != null || IsCol.SelectedValue != null ||
                       InCol.SelectedValue != null || Model.SelectedValue != null || Radius.SelectedValue != null || Width.SelectedValue != null || Spikes.SelectedValue != null)
                        filtered = FilterMainList(MainList, (string)Radius.SelectedValue, (string)Width.SelectedValue, (string)Height.SelectedValue, (string)Manufacturer.SelectedValue, (string)Model.SelectedValue, (string)IsCol.SelectedValue, (string)InCol.SelectedValue, (string)Gruz.SelectedValue, (string)RFT.SelectedValue, (string)Spikes.SelectedValue, (string)tb.SelectedValue);
                    else
                        filtered = FilterMainList(MainList, null, null, null, null, null, null, null, null, null, null, (string)tb.SelectedValue);
                    Details.ItemsSource = filtered;
                }
            }
            else
            {
                if (Height.SelectedValue != null || Gruz.SelectedValue != null || Manufacturer.SelectedValue != null || RFT.SelectedValue != null || IsCol.SelectedValue != null ||
                        InCol.SelectedValue != null || Model.SelectedValue != null || Radius.SelectedValue != null || Width.SelectedValue != null || Spikes.SelectedValue != null)
                    filtered = FilterMainList(MainList, (string)Radius.SelectedValue, (string)Width.SelectedValue, (string)Height.SelectedValue, (string)Manufacturer.SelectedValue, (string)Model.SelectedValue, (string)IsCol.SelectedValue, (string)InCol.SelectedValue, (string)Gruz.SelectedValue, (string)RFT.SelectedValue, (string)Spikes.SelectedValue, (string)tb.SelectedValue);
                else
                    Details.ItemsSource = tmpList;
            }
        }

        private void Manufacturer_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var tb = sender as ComboBox;
            var prods = MainList;//Details.ItemsSource as List<BdProducts>;
            if (tb.SelectedValue != null)
            {
                if (prods.Count > 0)
                {
                    if (Height.SelectedValue != null || Season.SelectedValue != null || Gruz.SelectedValue != null || RFT.SelectedValue != null || IsCol.SelectedValue != null ||
                       InCol.SelectedValue != null || Model.SelectedValue != null || Radius.SelectedValue != null || Width.SelectedValue != null || Spikes.SelectedValue != null)
                        filtered = FilterMainList(MainList, (string)Radius.SelectedValue, (string)Width.SelectedValue, (string)Height.SelectedValue, (string)tb.SelectedValue, (string)Model.SelectedValue, (string)IsCol.SelectedValue, (string)InCol.SelectedValue, (string)Gruz.SelectedValue, (string)RFT.SelectedValue, (string)Spikes.SelectedValue, (string)Season.SelectedValue);
                    else
                        filtered = FilterMainList(MainList, (string)Radius.SelectedValue, (string)Width.SelectedValue, (string)Height.SelectedValue, (string)tb.SelectedValue, (string)Model.SelectedValue, (string)IsCol.SelectedValue, (string)InCol.SelectedValue, (string)Gruz.SelectedValue, (string)RFT.SelectedValue, (string)Spikes.SelectedValue, (string)Season.SelectedValue);
                    Details.ItemsSource = filtered;
                }
            }
            else
            {
                if (Height.SelectedValue != null || Season.SelectedValue != null || Gruz.SelectedValue != null || RFT.SelectedValue != null || IsCol.SelectedValue != null ||
                        InCol.SelectedValue != null || Model.SelectedValue != null || Radius.SelectedValue != null || Width.SelectedValue != null || Spikes.SelectedValue != null)
                    filtered = FilterMainList(MainList, (string)Radius.SelectedValue, (string)Width.SelectedValue, (string)Height.SelectedValue, (string)tb.SelectedValue, (string)Model.SelectedValue, (string)IsCol.SelectedValue, (string)InCol.SelectedValue, (string)Gruz.SelectedValue, (string)RFT.SelectedValue, (string)Spikes.SelectedValue, (string)Season.SelectedValue);
                else
                    Details.ItemsSource = tmpList;
            }
        }

        private void Model_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var tb = sender as ComboBox;
            var prods = MainList;//Details.ItemsSource as List<BdProducts>;
            if (tb.SelectedValue != null)
            {
                if (prods.Count > 0)
                {
                    if (Height.SelectedValue != null || Season.SelectedValue != null || Manufacturer.SelectedValue != null || RFT.SelectedValue != null || IsCol.SelectedValue != null ||
                       InCol.SelectedValue != null || Gruz.SelectedValue != null || Radius.SelectedValue != null || Width.SelectedValue != null || Spikes.SelectedValue != null)
                        filtered = FilterMainList(MainList, (string)Radius.SelectedValue, (string)Width.SelectedValue, (string)Height.SelectedValue, (string)Manufacturer.SelectedValue, (string)tb.SelectedValue, (string)IsCol.SelectedValue, (string)InCol.SelectedValue, (string)Gruz.SelectedValue, (string)RFT.SelectedValue, (string)Spikes.SelectedValue, (string)Season.SelectedValue);
                    else
                        filtered = FilterMainList(MainList, (string)Radius.SelectedValue, (string)Width.SelectedValue, (string)Height.SelectedValue, (string)Manufacturer.SelectedValue, (string)tb.SelectedValue, (string)IsCol.SelectedValue, (string)InCol.SelectedValue, (string)Gruz.SelectedValue, (string)RFT.SelectedValue, (string)Spikes.SelectedValue, (string)Season.SelectedValue);
                    Details.ItemsSource = filtered;
                }
            }
            else
            {
                if (Height.SelectedValue != null || Season.SelectedValue != null || Manufacturer.SelectedValue != null || RFT.SelectedValue != null || IsCol.SelectedValue != null ||
                       InCol.SelectedValue != null || Gruz.SelectedValue != null || Radius.SelectedValue != null || Width.SelectedValue != null || Spikes.SelectedValue != null)
                    filtered = FilterMainList(MainList, (string)Radius.SelectedValue, (string)Width.SelectedValue, (string)Height.SelectedValue, (string)Manufacturer.SelectedValue, (string)tb.SelectedValue, (string)IsCol.SelectedValue, (string)InCol.SelectedValue, (string)Gruz.SelectedValue, (string)RFT.SelectedValue, (string)Spikes.SelectedValue, (string)Season.SelectedValue);
                else
                    Details.ItemsSource = tmpList;
            }
        }

        private void RFT_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var tb = sender as ComboBox;
            var prods = MainList;//Details.ItemsSource as List<BdProducts>;
            if (tb.SelectedValue != null)
            {
                if (prods.Count > 0)
                {
                    if (Height.SelectedValue != null || Season.SelectedValue != null || Manufacturer.SelectedValue != null || Gruz.SelectedValue != null || IsCol.SelectedValue != null ||
                        InCol.SelectedValue != null || Model.SelectedValue != null || Radius.SelectedValue != null || Width.SelectedValue != null || Spikes.SelectedValue != null)
                        filtered = FilterMainList(MainList, (string)Radius.SelectedValue, (string)Width.SelectedValue, (string)Height.SelectedValue, (string)Manufacturer.SelectedValue, (string)Model.SelectedValue, (string)IsCol.SelectedValue, (string)InCol.SelectedValue, (string)Gruz.SelectedValue, (string)tb.SelectedValue, (string)Spikes.SelectedValue, (string)Season.SelectedValue);
                    else
                        filtered = FilterMainList(MainList, (string)Radius.SelectedValue, (string)Width.SelectedValue, (string)Height.SelectedValue, (string)Manufacturer.SelectedValue, (string)Model.SelectedValue, (string)IsCol.SelectedValue, (string)InCol.SelectedValue, (string)Gruz.SelectedValue, (string)tb.SelectedValue, (string)Spikes.SelectedValue, (string)Season.SelectedValue);
                    Details.ItemsSource = filtered;
                }
            }
            else
            {
                if (Height.SelectedValue != null || Season.SelectedValue != null || Manufacturer.SelectedValue != null || Gruz.SelectedValue != null || IsCol.SelectedValue != null ||
                       InCol.SelectedValue != null || Model.SelectedValue != null || Radius.SelectedValue != null || Width.SelectedValue != null || Spikes.SelectedValue != null)
                    filtered = FilterMainList(MainList, (string)Radius.SelectedValue, (string)Width.SelectedValue, (string)Height.SelectedValue, (string)Manufacturer.SelectedValue, (string)Model.SelectedValue, (string)IsCol.SelectedValue, (string)InCol.SelectedValue, (string)Gruz.SelectedValue, (string)tb.SelectedValue, (string)Spikes.SelectedValue, (string)Season.SelectedValue);
                else
                    Details.ItemsSource = tmpList;
            }
        }

        private void IsCol_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var tb = sender as ComboBox;
            var prods = MainList;//Details.ItemsSource as List<BdProducts>;
            if (tb.SelectedValue != null)
            {
                if (prods.Count > 0)
                {
                    if (Height.SelectedValue != null || Season.SelectedValue != null || Manufacturer.SelectedValue != null || RFT.SelectedValue != null || Gruz.SelectedValue != null ||
                        InCol.SelectedValue != null || Model.SelectedValue != null || Radius.SelectedValue != null || Width.SelectedValue != null || Spikes.SelectedValue != null)
                        filtered = FilterMainList(MainList, (string)Radius.SelectedValue, (string)Width.SelectedValue, (string)Height.SelectedValue, (string)Manufacturer.SelectedValue, (string)Model.SelectedValue, (string)tb.SelectedValue, (string)InCol.SelectedValue, (string)Gruz.SelectedValue, (string)RFT.SelectedValue, (string)Spikes.SelectedValue, (string)Season.SelectedValue);
                    else
                        filtered = FilterMainList(MainList, (string)Radius.SelectedValue, (string)Width.SelectedValue, (string)Height.SelectedValue, (string)Manufacturer.SelectedValue, (string)Model.SelectedValue, (string)tb.SelectedValue, (string)InCol.SelectedValue, (string)Gruz.SelectedValue, (string)RFT.SelectedValue, (string)Spikes.SelectedValue, (string)Season.SelectedValue);
                    Details.ItemsSource = filtered;
                }
            }
            else
            {
                if (Height.SelectedValue != null || Season.SelectedValue != null || Manufacturer.SelectedValue != null || RFT.SelectedValue != null || Gruz.SelectedValue != null ||
                        InCol.SelectedValue != null || Model.SelectedValue != null || Radius.SelectedValue != null || Width.SelectedValue != null || Spikes.SelectedValue != null)
                    filtered = FilterMainList(MainList, (string)Radius.SelectedValue, (string)Width.SelectedValue, (string)Height.SelectedValue, (string)Manufacturer.SelectedValue, (string)Model.SelectedValue, (string)tb.SelectedValue, (string)InCol.SelectedValue, (string)Gruz.SelectedValue, (string)RFT.SelectedValue, (string)Spikes.SelectedValue, (string)Season.SelectedValue);
                else
                    Details.ItemsSource = tmpList;
            }
        }

        private void InCol_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var tb = sender as ComboBox;
            var prods = MainList;//Details.ItemsSource as List<BdProducts>;
            if (tb.SelectedValue != null)
            {
                if (prods.Count > 0)
                {
                    if (Height.SelectedValue != null || Season.SelectedValue != null || Manufacturer.SelectedValue != null || RFT.SelectedValue != null || IsCol.SelectedValue != null ||
                        Gruz.SelectedValue != null || Model.SelectedValue != null || Radius.SelectedValue != null || Width.SelectedValue != null || Spikes.SelectedValue != null)
                        filtered = FilterMainList(MainList, (string)Radius.SelectedValue, (string)Width.SelectedValue, (string)Height.SelectedValue, (string)Manufacturer.SelectedValue, (string)Model.SelectedValue, (string)IsCol.SelectedValue, (string)tb.SelectedValue, (string)Gruz.SelectedValue, (string)RFT.SelectedValue, (string)Spikes.SelectedValue, (string)Season.SelectedValue);
                    else
                        filtered = FilterMainList(MainList, (string)Radius.SelectedValue, (string)Width.SelectedValue, (string)Height.SelectedValue, (string)Manufacturer.SelectedValue, (string)Model.SelectedValue, (string)IsCol.SelectedValue, (string)tb.SelectedValue, (string)Gruz.SelectedValue, (string)RFT.SelectedValue, (string)Spikes.SelectedValue, (string)Season.SelectedValue);
                    Details.ItemsSource = filtered;
                }
            }
            else
            {
                if (Height.SelectedValue != null || Season.SelectedValue != null || Manufacturer.SelectedValue != null || RFT.SelectedValue != null || IsCol.SelectedValue != null ||
                       Gruz.SelectedValue != null || Model.SelectedValue != null || Radius.SelectedValue != null || Width.SelectedValue != null || Spikes.SelectedValue != null)
                    filtered = FilterMainList(MainList, (string)Radius.SelectedValue, (string)Width.SelectedValue, (string)Height.SelectedValue, (string)Manufacturer.SelectedValue, (string)Model.SelectedValue, (string)IsCol.SelectedValue, (string)tb.SelectedValue, (string)Gruz.SelectedValue, (string)RFT.SelectedValue, (string)Spikes.SelectedValue, (string)Season.SelectedValue);
                else
                    Details.ItemsSource = tmpList;
            }
        }

        private void Gruz_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var tb = sender as ComboBox;
            var prods = MainList;//Details.ItemsSource as List<BdProducts>;
            if (tb.SelectedValue != null)
            {
                if (prods.Count > 0)
                {
                    if (Height.SelectedValue != null || Season.SelectedValue != null || Manufacturer.SelectedValue != null || RFT.SelectedValue != null || IsCol.SelectedValue != null ||
                       InCol.SelectedValue != null || Model.SelectedValue != null || Radius.SelectedValue != null || Width.SelectedValue != null || Spikes.SelectedValue != null)
                        filtered = FilterMainList(MainList, (string)Radius.SelectedValue, (string)Width.SelectedValue, (string)Height.SelectedValue, (string)Manufacturer.SelectedValue, (string)Model.SelectedValue, (string)IsCol.SelectedValue, (string)InCol.SelectedValue, (string)tb.SelectedValue, (string)RFT.SelectedValue, (string)Spikes.SelectedValue, (string)Season.SelectedValue);
                    else
                        filtered = FilterMainList(MainList, (string)Radius.SelectedValue, (string)Width.SelectedValue, (string)Height.SelectedValue, (string)Manufacturer.SelectedValue, (string)Model.SelectedValue, (string)IsCol.SelectedValue, (string)InCol.SelectedValue, (string)tb.SelectedValue, (string)RFT.SelectedValue, (string)Spikes.SelectedValue, (string)Season.SelectedValue);
                    Details.ItemsSource = filtered;
                }
            }
            else
            {
                if (Height.SelectedValue != null || Season.SelectedValue != null || Manufacturer.SelectedValue != null || RFT.SelectedValue != null || IsCol.SelectedValue != null ||
                       InCol.SelectedValue != null || Model.SelectedValue != null || Radius.SelectedValue != null || Width.SelectedValue != null || Spikes.SelectedValue != null)
                    filtered = FilterMainList(MainList, (string)Radius.SelectedValue, (string)Width.SelectedValue, (string)Height.SelectedValue, (string)Manufacturer.SelectedValue, (string)Model.SelectedValue, (string)IsCol.SelectedValue, (string)InCol.SelectedValue, (string)tb.SelectedValue, (string)RFT.SelectedValue, (string)Spikes.SelectedValue, (string)Season.SelectedValue);
                else
                    Details.ItemsSource = tmpList;
            }
        }

        private void ResetFilters_Click(object sender, RoutedEventArgs e)
        {
            Details.ItemsSource = tmpList;
            using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
            {
                var widths = db.products.Select(s => s.Width.ToString()).Distinct().ToList();
                Width.ItemsSource = widths;
                var heights = db.products.Select(s => s.Height.ToString()).Distinct().ToList();
                Height.ItemsSource = heights;
                var radiuses = db.products.Select(s => s.Radius).Distinct().ToList();
                Radius.ItemsSource = radiuses;
            }
            foreach (var tb in tbFilters) {
                switch (tb.Name)
                {
                    case "Width":
                        Width.SelectionChanged -= Width_SelectionChanged;
                        Width.SelectedValue = -1;
                        Width.SelectionChanged += Width_SelectionChanged;
                        break;
                    case "Height":
                        Height.SelectionChanged -= Height_SelectionChanged;
                        Height.SelectedValue = -1;
                        Height.SelectionChanged += Height_SelectionChanged;
                        break;
                    case "Season":
                        Season.SelectionChanged -= Season_SelectionChanged;
                        Season.SelectedValue = -1;
                        Season.SelectionChanged += Season_SelectionChanged;
                        break;
                    case "RFT":
                        RFT.SelectionChanged -= RFT_SelectionChanged;
                        RFT.SelectedValue = -1;
                        RFT.SelectionChanged += RFT_SelectionChanged;
                        break;
                    case "IsCol":
                        IsCol.SelectionChanged -= IsCol_SelectionChanged;
                        IsCol.SelectedValue = -1;
                        IsCol.SelectionChanged += IsCol_SelectionChanged;
                        break;
                    case "InCol":
                        InCol.SelectionChanged -= InCol_SelectionChanged;
                        InCol.SelectedValue = -1;
                        InCol.SelectionChanged += InCol_SelectionChanged;
                        break;
                    case "Manufacturer":
                        Manufacturer.SelectionChanged -= Manufacturer_SelectionChanged;
                        Manufacturer.SelectedValue = -1;
                        Manufacturer.SelectionChanged += Manufacturer_SelectionChanged;
                        break;
                    case "Model":
                        Model.SelectionChanged -= Model_SelectionChanged;
                        Model.SelectedValue = -1;
                        Model.SelectionChanged += Model_SelectionChanged;
                        break;
                    case "Gruz":
                        Gruz.SelectionChanged -= Gruz_SelectionChanged;
                        Gruz.SelectedValue = -1;
                        Gruz.SelectionChanged += Gruz_SelectionChanged;
                        break;
                    case "Radius":
                        Radius.SelectionChanged -= Radius_SelectionChanged;
                        Radius.SelectedValue = -1;
                        Radius.SelectionChanged += Radius_SelectionChanged;
                        break;
                    case "Spikes":
                        Spikes.SelectionChanged -= Spikes_SelectionChanged;
                        Spikes.SelectedValue = -1;
                        Spikes.SelectionChanged += Spikes_SelectionChanged;
                        break;
                }
            }
        }

        private void Spikes_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var tb = sender as ComboBox;
            var prods = MainList;//Details.ItemsSource as List<BdProducts>;
            if (tb.SelectedValue != null)
            {
                if (prods.Count > 0)
                {
                    if (Height.SelectedValue != null || Gruz.SelectedValue != null || Manufacturer.SelectedValue != null || RFT.SelectedValue != null || IsCol.SelectedValue != null ||
                       InCol.SelectedValue != null || Model.SelectedValue != null || Radius.SelectedValue != null || Width.SelectedValue != null || Season.SelectedValue != null)
                        filtered = FilterMainList(MainList, (string)Radius.SelectedValue, (string)Width.SelectedValue, (string)Height.SelectedValue, (string)Manufacturer.SelectedValue, (string)Model.SelectedValue, (string)IsCol.SelectedValue, (string)InCol.SelectedValue, (string)Gruz.SelectedValue, (string)RFT.SelectedValue, (string)tb.SelectedValue, (string)Season.SelectedValue);
                    else
                        filtered = FilterMainList(MainList, (string)Radius.SelectedValue, (string)Width.SelectedValue, (string)Height.SelectedValue, (string)Manufacturer.SelectedValue, (string)Model.SelectedValue, (string)IsCol.SelectedValue, (string)InCol.SelectedValue, (string)Gruz.SelectedValue, (string)RFT.SelectedValue, (string)tb.SelectedValue, (string)Season.SelectedValue);
                    Details.ItemsSource = filtered;
                }
            }
            else
            {
                if (Height.SelectedValue != null || Gruz.SelectedValue != null || Manufacturer.SelectedValue != null || RFT.SelectedValue != null || IsCol.SelectedValue != null ||
                       InCol.SelectedValue != null || Model.SelectedValue != null || Radius.SelectedValue != null || Width.SelectedValue != null || Season.SelectedValue != null)
                    filtered = FilterMainList(MainList, (string)Radius.SelectedValue, (string)Width.SelectedValue, (string)Height.SelectedValue, (string)Manufacturer.SelectedValue, (string)Model.SelectedValue, (string)IsCol.SelectedValue, (string)InCol.SelectedValue, (string)Gruz.SelectedValue, (string)RFT.SelectedValue, (string)tb.SelectedValue, (string)Season.SelectedValue);
                else
                    Details.ItemsSource = tmpList;
            }
        }

        private void Refresh_Click(object sender, RoutedEventArgs e)
        {
            MainList = GetItemSource();
            Details.Items.Refresh();
            FillLoadControls();
            ProductTV.UpdateLayout();
        }
    }
}
