﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
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
    /// Логика взаимодействия для CapitalizeProduct.xaml
    /// </summary>
    public partial class CapitalizeProduct : Window
    {
        private List<product> prods;
        private List<PComboBox> list;
        private List<BdProducts> _mainList;
        private MainWindow _mainWnd;
        private Logger log;
        public CapitalizeProduct(MainWindow main, List<PComboBox> listP, List<BdProducts> bdp)
        {
            InitializeComponent();
            _mainWnd = main;
            list = listP;
            _mainList = bdp;
            log = LogManager.GetCurrentClassLogger();
            LoadControls();
            prods = new List<product>();
            
        }

        private void LoadControls()
        {
            try
            {
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    var categories = db.categories.ToList();
                    Categories.ItemsSource = categories;
                    Categories.SelectedValuePath = "CategoryId";
                    Categories.DisplayMemberPath = "CategoryName";

                    CategoriesMove.ItemsSource = categories;
                    CategoriesMove.SelectedValuePath = "CategoryId";
                    CategoriesMove.DisplayMemberPath = "CategoryName";

                    CategoriesOff.ItemsSource = categories;
                    CategoriesOff.SelectedValuePath = "CategoryId";
                    CategoriesOff.DisplayMemberPath = "CategoryName";
                    //Categories.SelectedValue = 1;
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void Categories_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var watch = System.Diagnostics.Stopwatch.StartNew();
            var cmb = sender as ComboBox;
            try
            {
                int category = (int)cmb.SelectedValue;
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    var manufacturers = db.manufacturers.Where(w => w.CategoryId == category).ToList();
                    Manufacturers.ItemsSource = manufacturers;
                    Manufacturers.SelectedValuePath = "ManufacturerId";
                    Manufacturers.DisplayMemberPath = "ManufacturerName";
                    Manufacturers.IsEnabled = true;
                    var elapsedMs = watch.ElapsedMilliseconds;
                    decimal minutes = (decimal)elapsedMs / (1000 * 60);
                    decimal sec = (decimal)elapsedMs / 1000;
                    //#endif
                    var specifier = "G";
                    var culture = CultureInfo.CreateSpecificCulture("ru-RU");
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void Manufacturers_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var watch = System.Diagnostics.Stopwatch.StartNew();
            try
            {
                int category = (int)Categories.SelectedValue;
                int ManId = (int)Manufacturers.SelectedValue;
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    var models = db.models.Where(w => w.CategoryId == category && w.ManufacturerId == ManId).ToList();
                    string manufacturer = db.manufacturers.Where(w => w.ManufacturerId == ManId).Select(s => s.ManufacturerName).First();
                    var products = db.products.ToList();

                    Models.ItemsSource = models;
                    Models.SelectedValuePath = "ModelId";
                    Models.DisplayMemberPath = "ModelName";
                    Models.IsEnabled = true;
                    var elapsedMs = watch.ElapsedMilliseconds;
                    decimal minutes = (decimal)elapsedMs / (1000 * 60);
                    decimal sec = (decimal)elapsedMs / 1000;
                    //#endif
                    var specifier = "G";
                    var culture = CultureInfo.CreateSpecificCulture("ru-RU");
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void Models_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                int category = (int)Categories.SelectedValue;
                int ManId = (int)Manufacturers.SelectedValue;
                var cmb = sender as ComboBox;
                int modelId = (int)cmb.SelectedValue;
                string modelTxt = cmb.Text;
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    //string manufacturer = db.manufacturers.Where(w => w.ManufacturerId == ManId).Select(s => s.ManufacturerName).First();
                    //var products = db.products.ToList();
                    var models = db.models.Where(w => w.CategoryId == category && w.ManufacturerId == ManId).ToList();
                    var radiuses = _mainList.Where(w => w.ModelId == modelId).Select(s => s.Radius).Distinct().ToList();
                    var widths = _mainList.Where(w => w.ModelId == modelId).Select(s => s.Width).Distinct().ToList();
                    var heights = _mainList.Where(w => w.ModelId == modelId).Select(s => s.Height).Distinct().ToList();
                    Radius.ItemsSource = radiuses;
                    Width.ItemsSource = widths;
                    Height.ItemsSource = heights;
                    var store = db.storehouses.ToList();
                    Storehouses.ItemsSource = store;
                    Storehouses.SelectedValuePath = "StorehouseId";
                    Storehouses.DisplayMemberPath = "StorehouseName";
                    Storehouses.IsEnabled = true;
                }
                //GridItemSource(category, prodId);
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void GridItemSource(int category, int prodId) {
            var watch = System.Diagnostics.Stopwatch.StartNew();
            try
            {
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    List<GridCapitalize> lst = new List<GridCapitalize>();
                    var product = _mainList.Single(w => w.CategoryId == category && w.ProductId == prodId);
                    var prodQuant = db.productquantities.Where(w => w.ProductId == product.ProductId).ToList();
                    foreach (var pq in prodQuant)
                    {
                        GridCapitalize gc = new GridCapitalize()
                        {
                            ProdName = list.Single(s => s.ProductId == product.ProductId).ProductName,
                            StorehouseName = db.storehouses.Single(s => s.StorehouseId == pq.StorehouseId).StorehouseName,
                            TotalQuantity = (int)pq.Quantity
                        };
                        lst.Add(gc);
                    }

                    var elapsedMs = watch.ElapsedMilliseconds;
                    decimal minutes = (decimal)elapsedMs / (1000 * 60);
                    decimal sec = (decimal)elapsedMs / 1000;
                    //#endif
                    var specifier = "G";
                    var culture = CultureInfo.CreateSpecificCulture("ru-RU");
                    Quantities.ItemsSource = lst;
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void InBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (int.TryParse(ProductInQuantity.Text, out int tmp))
                {
                    int category = (int)Categories.SelectedValue;
                    int ManId = (int)Manufacturers.SelectedValue;
                    int quant = int.Parse(ProductInQuantity.Text);
                    int store = (int)Storehouses.SelectedValue;
                    int modelId = (int)Models.SelectedValue;
                    string rad = (string)Radius.SelectedValue;
                    int width = (int)Width.SelectedValue;
                    float height = (float)Height.SelectedValue;
                    var product = GetProduct(category, ManId, modelId, rad, width, height);
                    Thread t = new Thread(() => InProduct(category, ManId, modelId, rad, width, height, store, quant));
                    t.Start();
                    GridItemSource(category, product.ProductId);
                }
                else
                {
                    MessageBox.Show("Невозможно добавить указанный товар на склад, введено неверное количество!", "Информация", MessageBoxButton.OK);
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void InProduct(int category, int ManId, int modelId, string rad, int width, float height, int store, int quant)
        {
            try
            {
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    var product = GetProduct(category, ManId, modelId, rad, width, height);
                    if (db.productquantities.Any(s => s.StorehouseId == store && s.ProductId == product.ProductId))
                    {
                        var storehouse = db.productquantities.Single(s => s.StorehouseId == store && s.ProductId == product.ProductId);
                        storehouse.Quantity += quant;
                        db.Entry(storehouse).Property(p => p.Quantity).IsModified = true;
                        db.SaveChanges();
                        //GridItemSource(category, product.ProductId);
                        MessageBox.Show("Товар успешно оприходован!", "Информация", MessageBoxButton.OK);
                    }
                    else
                    {
                        productquantity pq = new productquantity()
                        {
                            StorehouseId = store,
                            ProductId = product.ProductId,
                            Quantity = quant
                        };
                        db.productquantities.Add(pq);
                        db.SaveChanges();
                        //GridItemSource(category, product.ProductId);
                        MessageBox.Show("Товар успешно добавлен на указанный склад!", "Информация", MessageBoxButton.OK);
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }
        
        private product GetProduct(int category, int ManId, int modelId, string rad, int width, float height) {
            var watch = System.Diagnostics.Stopwatch.StartNew();
            product product = new product();
            try
            {
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    product = db.products.Single(s => s.CategoryId == category && s.ManufacturerId == ManId && s.ModelId == modelId && s.Radius == rad && s.Width == width && s.Height == height);
                    watch.Stop();
                    var elapsedMs = watch.ElapsedMilliseconds;
                    decimal minutes = (decimal)elapsedMs / (1000 * 60);
                    decimal sec = (decimal)elapsedMs / 1000;
                    //#endif
                    var specifier = "G";
                    var culture = CultureInfo.CreateSpecificCulture("ru-RU");
                    
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                log.Trace(category + "/" + ManId + "/" + modelId + "/" + rad + "/" + width + "/" + height);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
            return product;
        }

        private void MoveBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (int.TryParse(ProductMoveQuantity.Text, out int tmp))
                {
                    int category = (int)CategoriesMove.SelectedValue;
                    int ManId = (int)ManufacturersMove.SelectedValue;
                    int quantMove = int.Parse(ProductMoveQuantity.Text);
                    int storeFrom = (int)StorehouseFrom.SelectedValue;
                    int storeTo = (int)StorehouseTo.SelectedValue;
                    int modelId = (int)ModelsMove.SelectedValue;
                    string rad = (string)RadiusMove.SelectedValue;
                    int width = (int)WidthMove.SelectedValue;
                    float height = (float)HeightMove.SelectedValue;
                    int quantFrom = int.Parse(QuantAtStore.Content.ToString());
                    var product = GetProduct(category, ManId, modelId, rad, width, height);
                    Thread t = new Thread(() => MoveProduct(category, ManId, quantMove, storeFrom, storeTo, modelId, rad, width, height, quantFrom));
                    t.Start();
                    Thread.Sleep(1000);
                    GridItemSource(category, product.ProductId);
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void MoveProduct(int category, int ManId, int quantMove, int storeFrom, int storeTo, int modelId, string rad, int width, float height, int quantFrom) {
            try
            {
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {

                    var diff = quantFrom - quantMove;
                    if (diff >= 0)
                    {
                        var product = GetProduct(category, ManId, modelId, rad, width, height);
                        if (db.productquantities.Any(s => s.StorehouseId == storeTo && s.ProductId == product.ProductId))
                        {
                            var storehouseTo = db.productquantities.Single(s => s.StorehouseId == storeTo && s.ProductId == product.ProductId);
                            storehouseTo.Quantity += quantMove;
                            db.Entry(storehouseTo).Property(p => p.Quantity).IsModified = true;
                            var storehouseFrom = db.productquantities.Single(s => s.ProductId == product.ProductId && s.StorehouseId == storeFrom);
                            storehouseFrom.Quantity -= quantMove;
                            db.Entry(storehouseFrom).Property(p => p.Quantity).IsModified = true;
                            db.SaveChanges();

                            MessageBox.Show("Товар успешно перенесен на указанный склад!", "Информация", MessageBoxButton.OK);
                        }
                        else
                        {
                            productquantity pq = new productquantity()
                            {
                                StorehouseId = storeTo,
                                ProductId = product.ProductId,
                                Quantity = quantMove
                            };
                            db.productquantities.Add(pq);
                            var storehouseFrom = db.productquantities.Single(s => s.ProductId == product.ProductId && s.StorehouseId == storeFrom);
                            storehouseFrom.Quantity -= quantMove;
                            db.Entry(storehouseFrom).Property(p => p.Quantity).IsModified = true;
                            db.SaveChanges();
                            // GridItemSource(category, product.ProductId);
                            MessageBox.Show("Товар успешно перенесен на указанный склад!", "Информация", MessageBoxButton.OK);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Невозможно перенести указанный товар со склада, введено неверное количество!", "Информация", MessageBoxButton.OK);
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void CategoriesMove_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var cmb = sender as ComboBox;
            try
            {
                int category = (int)cmb.SelectedValue;
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    var manufacturers = db.manufacturers.Where(w => w.CategoryId == category).ToList();
                    ManufacturersMove.ItemsSource = manufacturers;
                    ManufacturersMove.SelectedValuePath = "ManufacturerId";
                    ManufacturersMove.DisplayMemberPath = "ManufacturerName";
                    ManufacturersMove.IsEnabled = true;
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void ManufacturersMove_SelectionChanged_1(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                int category = (int)CategoriesMove.SelectedValue;
                int ManId = (int)ManufacturersMove.SelectedValue;
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    var models = db.models.Where(w => w.CategoryId == category && w.ManufacturerId == ManId).ToList();
                    //string manufacturer = db.manufacturers.Where(w=>w.ManufacturerId==ManId).Select(s=>s.ManufacturerName).First();
                    prods = db.products.Where(w => w.ManufacturerId == ManId).ToList();
                    ModelsMove.ItemsSource = models;
                    ModelsMove.SelectedValuePath = "ModelId";
                    ModelsMove.DisplayMemberPath = "ModelName";
                    ModelsMove.IsEnabled = true;
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void ModelsMove_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                int category = (int)CategoriesMove.SelectedValue;
                int ManId = (int)ManufacturersMove.SelectedValue;
                var cmb = sender as ComboBox;
                int modelId = (int)cmb.SelectedValue;
                string modelTxt = cmb.Text;
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    //string manufacturer = db.manufacturers.Where(w => w.ManufacturerId == ManId).Select(s => s.ManufacturerName).First();
                    //var products = db.products.ToList();
                    var models = db.models.Where(w => w.CategoryId == category && w.ManufacturerId == ManId).ToList();
                    var radiuses = _mainList.Where(w => w.ModelId == modelId).Select(s => s.Radius).Distinct().ToList();
                    var widths = _mainList.Where(w => w.ModelId == modelId).Select(s => s.Width).Distinct().ToList();
                    var heights = _mainList.Where(w => w.ModelId == modelId).Select(s => s.Height).Distinct().ToList();
                    RadiusMove.ItemsSource = radiuses;
                    WidthMove.ItemsSource = widths;
                    HeightMove.ItemsSource = heights;
                    prods = prods.Where(w => w.ModelId == modelId).ToList();
                    
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void StorehouseFrom_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var cmb = sender as ComboBox;
                int value = (int)cmb.SelectedValue;
                int category = (int)CategoriesMove.SelectedValue;
                int ManId = (int)ManufacturersMove.SelectedValue;
                int modelId = (int)ModelsMove.SelectedValue;
                string rad = (string)RadiusMove.SelectedValue;
                int width = (int)WidthMove.SelectedValue;
                float height = (float)HeightMove.SelectedValue;
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    int prodId = GetProduct(category, ManId, modelId, rad, width, height).ProductId;
                    QuantAtStore.Content = db.productquantities.Where(w => w.ProductId == prodId && w.StorehouseId == value).Select(s => s.Quantity.ToString()).First();
                    var storehouses = db.storehouses.Where(w => w.StorehouseId != value).ToList();
                    StorehouseTo.ItemsSource = storehouses;
                    StorehouseTo.SelectedValuePath = "StorehouseId";
                    StorehouseTo.DisplayMemberPath = "StorehouseName";
                    StorehouseTo.IsEnabled = true;
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void CategoriesOff_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var cmb = sender as ComboBox;
            try
            {
                int category = (int)cmb.SelectedValue;
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    var manufacturers = db.manufacturers.Where(w => w.CategoryId == category).ToList();
                    ManufacturersOff.ItemsSource = manufacturers;
                    ManufacturersOff.SelectedValuePath = "ManufacturerId";
                    ManufacturersOff.DisplayMemberPath = "ManufacturerName";
                    ManufacturersOff.IsEnabled = true;
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void ManufacturersOff_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //List<PComboBox> list = new List<PComboBox>();
            int category = (int)CategoriesOff.SelectedValue;
            try
            {
                int ManId = (int)ManufacturersOff.SelectedValue;
                //using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                //{
                //var models = db.models.Where(w => w.CategoryId == category && w.ManufacturerId == ManId).ToList();
                //string manufacturer = db.manufacturers.Where(w => w.ManufacturerId == ManId).Select(s => s.ManufacturerName).First();
                //var products = db.products.ToList();
                //foreach (var model in models)
                //{
                //    var range = products.Where(w => w.ModelId == model.ModelId && w.ManufacturerId == ManId).Select(s => new PComboBox()
                //    {
                //        ProductId = s.ProductId,
                //        ProductName = @"" + manufacturer + " " + model.ModelName + " " + s.Width + " / " + s.Height + " / " + s.Radius
                //    }).ToList();
                //    list.AddRange(range);
                //}
                // list = 
                ModelsOff.ItemsSource = list;
                ModelsOff.SelectedValuePath = "ProductId";
                ModelsOff.DisplayMemberPath = "ProductName";
                ModelsOff.IsEnabled = true;
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
            //}
        }

        private void ModelsOff_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var cmb = sender as ComboBox;
            try
            {
                int prodId = (int)cmb.SelectedValue;
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    var tmp = db.productquantities.Where(w => w.ProductId == prodId && w.Quantity > 0).ToList();
                    var storeTmp = db.storehouses.ToList();
                    var storehouses = storeTmp.Where(w => tmp.Any(a => a.StorehouseId == w.StorehouseId)).ToList();
                    StorehouseOffFrom.ItemsSource = storehouses;
                    StorehouseOffFrom.SelectedValuePath = "StorehouseId";
                    StorehouseOffFrom.DisplayMemberPath = "StorehouseName";
                    StorehouseOffFrom.IsEnabled = true;
                    ProductOffPrice.Text = db.products.Where(w => w.ProductId == prodId).Select(s => s.Price.ToString()).First();
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void OffBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int prodId = (int)ModelsOff.SelectedValue;
                if (int.TryParse(ProductOffQuantity.Text, out var tmp))
                {
                    int quant = int.Parse(ProductOffQuantity.Text);
                    int store = (int)StorehouseOffFrom.SelectedValue;
                    CultureInfo provider = CultureInfo.InvariantCulture;
                    using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                    {
                        var storeHouse = db.productquantities.Single(w => w.ProductId == prodId && w.StorehouseId == store);
                        var sNum = db.operations.ToList();
                        long number = 0;
                        if (sNum.Count == 0)
                            number = 1;
                        else
                            number = (long)sNum[sNum.Count - 1].SaleNumber + 1;
                        var prod = db.products.Single(s => s.ProductId == prodId);
                        var date = DateTime.Now.ToString("dd-MM-yyyy");
                        var time = DateTime.Now.ToString("hh:mm:ss");
                        var diff = storeHouse.Quantity - quant;
                        if (diff >= 0)
                        {
                            operation oper = new operation()
                            {
                                OperationDate = DateTime.ParseExact(date, "dd-MM-yyyy", provider),
                                OperationTime = DateTime.ParseExact(time, "hh:mm:ss", provider),
                                Price = decimal.Parse(ProductOffPrice.Text, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture) * (decimal)quant,
                                Count = quant,
                                ProductId = prodId,
                                //ServiceId = item.ServiceId,
                                SaleNumber = number,
                                Comment = OffComment.Text == string.Empty ? "" : OffComment.Text,
                                OperationType = "Списание товара",
                                PayType = "Списание товара",
                                Storehouse = db.storehouses.Where(w => w.StorehouseId == store).Select(s => s.StorehouseName).First(),
                                CardPay = "Нет"
                                /*
                                PayType = item.PayType,
                                OperationType = item.OperationType,
                                Comment = item.Comment,
                                Storehouse = storehouseName,
                                CardPay = item.CardPayed
                                 */
                            };
                            db.operations.Add(oper);
                            storeHouse.Quantity -= quant;
                            db.Entry(storeHouse).Property(p => p.Quantity).IsModified = true;
                            db.SaveChanges();
                            MessageBox.Show("Товар успешно списан!", "Информация", MessageBoxButton.OK);
                        }
                        else
                            MessageBox.Show("На выбранном складе недостаточное количество товара для осуществления операции списания!", "Информация", MessageBoxButton.OK);
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void WidthMove_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var cmb = sender as ComboBox;
            try
            {
                int width = (int)cmb.SelectedValue;
                prods = prods.Where(w => w.Width == width).ToList();
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void HeightMove_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var cmb = sender as ComboBox;
            try
            {
                float height = (float)cmb.SelectedValue;
                prods = prods.Where(w => w.Height == height).ToList();
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void RadiusMove_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var cmb = sender as ComboBox;
            try
            {
                string rad = (string)cmb.SelectedValue;
                prods = prods.Where(w => w.Radius == rad).ToList();
                if (prods.Count == 1)
                {
                    var product = prods[0];
                    using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                    {
                        var store = db.storehouses.ToList();
                        var tmp = db.productquantities.Where(w => w.ProductId == product.ProductId && w.Quantity > 0).ToList();
                        var storeTmp = db.storehouses.ToList();
                        var storehouses = storeTmp.Where(w => tmp.Any(a => a.StorehouseId == w.StorehouseId)).ToList();
                        StorehouseFrom.ItemsSource = storehouses;
                        StorehouseFrom.SelectedValuePath = "StorehouseId";
                        StorehouseFrom.DisplayMemberPath = "StorehouseName";
                        StorehouseFrom.IsEnabled = true;
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void Radius_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try {
                int category = (int)Categories.SelectedValue;
                int ManId = (int)Manufacturers.SelectedValue;
                int modelId = (int)Models.SelectedValue;
                string rad = (string)Radius.SelectedValue;
                int width = (int)Width.SelectedValue;
                float height = (float)Height.SelectedValue;
                var product = GetProduct(category, ManId, modelId, rad, width, height);
                GridItemSource(category, product.ProductId);
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }
    }
}
