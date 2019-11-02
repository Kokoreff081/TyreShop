using System;
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
                string model = "", manufacturer = "";
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    var product = GetProduct(category, ManId, modelId, rad, width, height);
                    model = db.models.Single(s => s.ModelId == modelId).ModelName;
                    manufacturer = db.manufacturers.Single(s => s.ManufacturerId == ManId).ManufacturerName;
                    if (db.productquantities.Any(s => s.StorehouseId == store && s.ProductId == product.ProductId))
                    {
                        var storehouse = db.productquantities.Single(s => s.StorehouseId == store && s.ProductId == product.ProductId);
                        storehouse.Quantity += quant;
                        db.Entry(storehouse).Property(p => p.Quantity).IsModified = true;
                        if (storehouse.Quantity > 0)
                            product.ProdStatus = true;
                        else
                            product.ProdStatus = false;
                        if (db.products.Any(a => a.ProductId == product.ProductId))
                        {
                            var prod = db.products.Single(s => s.ProductId == product.ProductId);
                            prod.ProdStatus = product.ProdStatus;
                            db.Entry(prod).Property(p => p.ProdStatus).IsModified = true;
                        }
                        
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
                using (u0324292_mainEntities db2 = new u0324292_mainEntities())
                {
                    var product = GetProduct(category, ManId, modelId, rad, width, height);
                    var id = uint.Parse(product.ProdNumber);
                    if (db2.shop_product.Any(a => a.product_id == id))
                    {
                        var siteProd = db2.shop_product.Single(a => a.product_id == id);
                        siteProd.quantity += quant;
                        db2.Entry(siteProd).Property(p => p.quantity).IsModified = true;
                        if (siteProd.quantity >= 0)
                            db2.SaveChanges();
                        if (siteProd.quantity >= 0 && siteProd.stock_status_id == 8)
                        {
                            siteProd.stock_status_id = 7;
                            db2.Entry(siteProd).Property(p => p.stock_status_id).IsModified = true;
                            db2.SaveChangesAsync();
                        }
                    }
                    else
                    {
                        int shopStat = 0;
                        if (quant > 0)//больше нуля - продаем
                            shopStat = 7;
                        else//в любом другом случае предзаказ
                            shopStat = 8;
                        shop_product sp = new shop_product()
                        {
                            product_id = id,//айдишник продукта в сайтовой базе
                            model = model,//наименование модели
                            quantity = quant,//количество в наличии
                            stock_status_id = shopStat,//статус
                            manufacturer_id = db2.shop_manufacturer.Single(s => s.name == manufacturer).manufacturer_id,//получение и присвоение айдишника производителя
                            price = (decimal)product.Price,//цена
                            status = true,
                            //subtract = true,

                        };
                        List<shop_product_attribute> lst = new List<shop_product_attribute>()//список атрибутов
                        {
                            new shop_product_attribute(){ product_id = int.Parse(product.Articul), attribute_id = 12, text = width.ToString()},//ширина
                            new shop_product_attribute(){ product_id = int.Parse(product.Articul), attribute_id = 13, text = height.ToString()},//высота
                            new shop_product_attribute(){ product_id = int.Parse(product.Articul), attribute_id = 14, text = product.Radius},//радиус
                            new shop_product_attribute(){ product_id = int.Parse(product.Articul), attribute_id = 15, text = product.Season},//сезон, у меня - зима или лето
                            new shop_product_attribute(){ product_id = int.Parse(product.Articul), attribute_id = 16, text = product.InCol},//ИН
                            new shop_product_attribute(){ product_id = int.Parse(product.Articul), attribute_id = 17, text = product.IsCol},//ИС
                            new shop_product_attribute(){ product_id = int.Parse(product.Articul), attribute_id = 18, text = product.RFT},//РанФлэт
                            new shop_product_attribute(){ product_id = int.Parse(product.Articul), attribute_id = 19, text = product.Gruz},//Грузовой
                            new shop_product_attribute(){ product_id = int.Parse(product.Articul), attribute_id = 20, text = product.Spikes},//Шипы
                        };
                        shop_product_to_category sptc = new shop_product_to_category()
                        {
                            category_id = 1,
                            product_id = id
                        };
                        shop_product_to_layout sptl = new shop_product_to_layout()
                        {
                            layout_id = 0,
                            store_id = 0,
                            product_id = id
                        };
                        shop_product_to_store spts = new shop_product_to_store()
                        {
                            product_id = id,
                            store_id = 0
                        };
                        string seasonToMetaTitle = "";
                        switch (product.Season)
                        {
                            case "Лето":
                                seasonToMetaTitle = "Летние";
                                break;
                            case "Зима":
                                seasonToMetaTitle = "Зимние";
                                break;
                        }
                        string productName = manufacturer + " " + model + " " + width + "/" + height + product.Radius + " " + product.InCol + product.IsCol + " " + product.RFT + product.Spikes;
                        shop_product_description spd = new shop_product_description()
                        {
                            product_id = id,
                            language_id = 1,
                            name = productName,
                            description = "",
                            tag = width + "," + height + "," + product.Radius,
                            meta_title = seasonToMetaTitle + " шины " + productName + ". Магазин автошин TireShop",
                            meta_description = seasonToMetaTitle + " шины " + productName + " по " + product.Price * 4 + " руб/шт. Доставка по СПб и в другие регионы",
                            meta_keyword = ""
                        };
                        shop_url_alias sua = new shop_url_alias()
                        {
                            query = "product_id=" + id,
                            keyword = manufacturer + "_" + id
                        };
                        db2.shop_product.Add(sp);//добавили сам продукт
                        foreach (var item in lst)
                        {//добавляем все атрибуты шин
                            db2.shop_product_attribute.Add(item);
                        }
                        db2.shop_product_to_category.Add(sptc);
                        db2.shop_product_to_layout.Add(sptl);
                        db2.shop_product_to_store.Add(spts);
                        db2.shop_product_description.Add(spd);
                        db2.shop_url_alias.Add(sua);
                        db2.SaveChanges();
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
                product prod = new product();
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
                        prod = db.products.Single(s => s.ProductId == prodId);
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
                            if (storeHouse.Quantity > 0)
                                prod.ProdStatus = true;
                            else
                                prod.ProdStatus = false;
                            db.Entry(prod).Property(p => p.ProdStatus).IsModified = true;
                            db.SaveChanges();
                            MessageBox.Show("Товар успешно списан!", "Информация", MessageBoxButton.OK);
                        }
                        else
                            MessageBox.Show("На выбранном складе недостаточное количество товара для осуществления операции списания!", "Информация", MessageBoxButton.OK);
                    }
                    using (u0324292_mainEntities db2 = new u0324292_mainEntities())
                    {
                        var id = int.Parse(prod.ProdNumber);
                        if (db2.shop_product.Any(a => a.product_id == id))
                        {
                            var siteProd = db2.shop_product.Single(a => a.product_id == id);
                            siteProd.quantity -= quant;
                            db2.Entry(siteProd).Property(p => p.quantity).IsModified = true;
                            if (siteProd.quantity >= 0)
                                db2.SaveChanges();
                            else
                                log.Error(siteProd.model + siteProd.quantity + siteProd.product_id + " количество не может быть отрицательным");
                            if (siteProd.quantity == 0 && siteProd.stock_status_id == 7)
                            {
                                siteProd.stock_status_id = 8;
                                db2.Entry(siteProd).Property(p => p.stock_status_id).IsModified = true;
                                db2.SaveChangesAsync();
                            }

                        }
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
