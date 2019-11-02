using Microsoft.Win32;
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
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.ComponentModel;
using System.Globalization;
using System.Data.Entity.Validation;
using NLog;

namespace Tyreshop
{
    /// <summary>
    /// Логика взаимодействия для AddProduct.xaml
    /// </summary>
    public partial class AddProduct : Window
    {
        private MainWindow _mainWnd;
        private BackgroundWorker backgroundWorker;
        public List<PComboBox> listP;
        private List<BdProducts> _MainList;
        private Logger log;
        public AddProduct(MainWindow mWnd, List<BdProducts> list, List<PComboBox> lpc)
        {
            InitializeComponent();
            _mainWnd = mWnd;
            _MainList = list;
            listP = lpc;
            log = LogManager.GetCurrentClassLogger();
            LoadControls();
        }

        private void LoadControls() {
            try
            {
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    var categories = db.categories.ToList();
                    Categories.ItemsSource = categories;
                    Categories.SelectedValuePath = "CategoryId";
                    Categories.DisplayMemberPath = "CategoryName";
                    Categories.SelectedValue = 1;

                    CategoriesInModels.ItemsSource = categories;
                    CategoriesInModels.SelectedValuePath = "CategoryId";
                    CategoriesInModels.DisplayMemberPath = "CategoryName";

                    Categories2.ItemsSource = categories;
                    Categories2.SelectedValuePath = "CategoryId";
                    Categories2.DisplayMemberPath = "CategoryName";

                    var rads = new List<string>() {
                    "R13","R14","R15","R16","R17","R18","R19","R20","R21","R22","R23","R24"
                };//db.products.Select(s => s.Radius).Distinct().ToList();
                    Radius.ItemsSource = rads;
                    RadiusEdit.ItemsSource = rads;

                    var widths = new List<string>() {
                    "31", "32", "33", "145","155","165","175","185","195","205","215","225","235","245","255","265","275","285","295","305","315","325","335"
                };//db.products.Select(s => s.Width).Distinct().ToList();
                    Width.ItemsSource = widths;
                    WidthEdit.ItemsSource = widths;

                    var heigths = new List<string>() {
                    "10.5", "12.5", "25","30","35","40","45","50","55","60","65","70","75","80","85","90","95","100",
                };//db.products.Select(s => s.Height).Distinct().ToList();

                    Height.ItemsSource = heigths;
                    HeightEdit.ItemsSource = heigths;

                    ModelsEdit.ItemsSource = listP;
                    ModelsEdit.SelectedValuePath = "ProductId";
                    ModelsEdit.DisplayMemberPath = "ProductName";

                    var stores = db.storehouses.ToList();
                    Stores.ItemsSource = stores;
                    Stores.SelectedValuePath = "StorehouseId";
                    Stores.DisplayMemberPath = "StorehouseName";
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
            var cmb = sender as ComboBox;
            try
            {
                int manufacturer = (int)cmb.SelectedValue;

                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    var models = db.models.Where(w => w.ManufacturerId == manufacturer).ToList();
                    Models.ItemsSource = models;
                    Models.SelectedValuePath = "ModelId";
                    Models.DisplayMemberPath = "ModelName";
                    Models.IsEnabled = true;
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void AddManufacturer_Click(object sender, RoutedEventArgs e)
        {
            int catId = (int)Categories.SelectedValue;
            string name = ManufacturerName.Text;
            if (name != string.Empty) {
                try
                {
                    using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                    {
                        manufacturer newMan = new manufacturer()
                        {
                            CategoryId = catId,
                            ManufacturerName = name
                        };
                        db.manufacturers.Add(newMan);
                        db.SaveChanges();
                        MessageBox.Show("Производитель успешно добавлен!", "Информация", MessageBoxButton.OK);
                    }
                }
                catch (Exception ex)
                {
                    log.Error(ex.Message + " \n" + ex.StackTrace);
                    MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
                }
            }
            else
            {
                MessageBox.Show("Введите наименование производителя!", "Информация", MessageBoxButton.OK);
            }
        }

        private void AddModel_Click(object sender, RoutedEventArgs e)
        {
            int catId = (int)CategoriesInModels.SelectedValue;
            int manId = (int)Manufacturers.SelectedValue;
            string name = ModelName.Text;
            if (name != string.Empty)
            {
                try
                {
                    using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                    {
                        model newModel = new model()
                        {
                            CategoryId = catId,
                            ManufacturerId = manId,
                            ModelName = name
                        };
                        db.models.Add(newModel);
                        db.SaveChanges();
                        MessageBox.Show("Модель успешно добавлена!", "Информация", MessageBoxButton.OK);
                    }
                }
                catch (Exception ex)
                {
                    log.Error(ex.Message + " \n" + ex.StackTrace);
                    MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
                }
            }
            else
            {
                MessageBox.Show("Введите наименование модели!", "Информация", MessageBoxButton.OK);
            }
        }

        private void AddProductBtn_Click(object sender, RoutedEventArgs e)
        {
            int catId = (int)Categories2.SelectedValue;
            int manId = (int)Manufacturers2.SelectedValue;
            int modelId = (int)Models.SelectedValue;
            int tmp = -1; float tmp3 = -1;
            decimal tmp2 = -1; 
            bool Flag = true;
            int width = 0, quant = 0; float height = 0;
            if(Quant.Text!=string.Empty)
                quant = int.Parse(Quant.Text);
            decimal price = 0, optPrice = 0, purPrice=0;
            string radius = "", art = "", season = "", inCol = "", isCol = "", gruz = "", rft = "", spike = "", country="", manufacturer = "", model="";
            if (Radius.Text != string.Empty)
                radius = Radius.Text;
            else
                Flag = false;
            if (int.TryParse(Width.Text, out tmp))
            {
                width = int.Parse(Width.Text);
            }
            else
                Flag = false;

            if (float.TryParse(Height.Text, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out tmp3))
            {
                height = float.Parse(Height.Text, CultureInfo.InvariantCulture);
            }
            else
                Flag = false;
            if (Season.Text != string.Empty)
            {
                season = Season.Text;
            }
            else
                Flag = false;
            if (Country.Text != string.Empty)
            {
                country = Country.Text;
            }
            else
                Flag = false;
            if (InCol.Text != string.Empty)
            {
                inCol = InCol.Text;
            }
            else
                Flag = false;
            if (IsCol.Text != string.Empty)
            {
                isCol = IsCol.Text;
            }
            else
                Flag = false;
            if (Gruz.Text != string.Empty)
            {
                gruz = Gruz.Text;
                if (gruz == "Да")
                {
                    //Radius.Text += "C";
                    radius += "C";
                }
            }
            else
                Flag = false;
            if (RFT.Text != string.Empty)
            {
                rft = RFT.Text;
            }
            else
                Flag = false;
            if (Articul.Text != string.Empty)
            {
                art = Articul.Text;
            }
            else
                Flag = false;
            if (Spikes.Text != string.Empty)
            {
                spike = Spikes.Text;
            }
            else
                Flag = false;
            if (decimal.TryParse(Price.Text, out tmp2))
            {
                price = decimal.Parse(Price.Text);
            }
            else
                Flag = false;
            if (decimal.TryParse(OptPrice.Text, out tmp2))
            {
                optPrice = decimal.Parse(OptPrice.Text);
            }
            else
                Flag = false;
            if (decimal.TryParse(PurchasePrice.Text, out tmp2))
            {
                purPrice = decimal.Parse(PurchasePrice.Text);
            }
            else
                Flag = false;
            if (Flag)
            {
                try
                {
                    using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                    {
                        if (!db.products.Any(a => a.CategoryId == catId && a.ManufacturerId == manId && a.ModelId == modelId && a.Radius == radius && a.ProdNumber == art && a.Width == width && a.Height == height
                         && a.InCol == inCol && a.IsCol == isCol && a.Gruz == gruz && a.OptPrice == optPrice && a.Price == price && a.PurchasePrice == purPrice && a.Country == country && a.RFT == rft && a.Season == season
                         && a.Spikes == spike))
                        {
                            product newProd = new product()
                            {
                                CategoryId = catId,
                                ManufacturerId = manId,
                                ModelId = modelId,
                                Radius = radius,
                                ProdNumber = art,
                                Gruz = gruz,
                                Height = height,
                                Width = width,
                                InCol = inCol,
                                IsCol = isCol,
                                OptPrice = optPrice,
                                Price = price,
                                RFT = rft,
                                Season = season,
                                Spikes = spike,
                                Country = country,
                                PurchasePrice = purPrice
                            };
                            db.products.Add(newProd);
                            db.SaveChanges();
                            manufacturer = db.manufacturers.Single(s => s.ManufacturerId == manId).ManufacturerName;
                            model = db.models.Single(s => s.ModelId == modelId).ModelName;
                            var prodId = newProd.ProductId;
                            if (Stores.SelectedValue != null)
                            {
                                int storeId = (int)Stores.SelectedValue;
                                if (Quant.Text != string.Empty)
                                {
                                    if (int.TryParse(Quant.Text, out var tmp4))
                                    {
                                        
                                        var pq = new productquantity()
                                        {
                                            ProductId = prodId,
                                            StorehouseId = storeId,
                                            Quantity = quant
                                        };
                                        if (quant > 0)
                                            newProd.ProdStatus = true;
                                        else
                                            newProd.ProdStatus = false;
                                        db.Entry(newProd).Property(p => p.ProdStatus).IsModified = true;
                                        db.productquantities.Add(pq);
                                        db.SaveChanges();
                                    }
                                }
                            }
                            Width.Text = "";
                            Height.Text = "";
                            Radius.Text = "";
                            InCol.Text = "";
                            IsCol.Text = "";
                            Country.Text = "";
                            Gruz.Text = "";
                            Spikes.Text = "";
                            RFT.Text = "";
                            Season.Text = "";
                            Price.Text = "";
                            PurchasePrice.Text = "";
                            OptPrice.Text = "";
                            //ModelsEdit.SelectionChanged -= ModelsEdit_SelectionChanged;
                            Models.SelectedValue = -1;
                            Articul.Text = "";
                            //ModelsEdit.SelectionChanged += ModelsEdit_SelectionChanged;
                            MessageBox.Show("Товар успешно добавлен!", "Информация", MessageBoxButton.OK);
                            using (u0324292_mainEntities db2 = new u0324292_mainEntities())
                            {
                                int shopStat = 0;
                                if (quant > 0)//больше нуля - продаем
                                    shopStat = 7;
                                else//в любом другом случае предзаказ
                                    shopStat = 8;
                                var id = uint.Parse(art);
                                shop_product sp = new shop_product()
                                {
                                    product_id = id,//айдишник продукта в сайтовой базе
                                    model = model,//наименование модели
                                    quantity = quant,//количество в наличии
                                    stock_status_id = shopStat,//статус
                                    manufacturer_id = db2.shop_manufacturer.Single(s => s.name == manufacturer).manufacturer_id,//получение и присвоение айдишника производителя
                                    price = price,//цена
                                    status = true,
                                    sku = "",
                                    upc = "",
                                    ean = "",
                                    jan = "",
                                    isbn = "",
                                    mpn = "",
                                    location = ""
                                    //subtract = true,
                                };
                                List<shop_product_attribute> lst = new List<shop_product_attribute>()//список атрибутов
                                {
                                    new shop_product_attribute(){ product_id = int.Parse(art), attribute_id = 12, text = width.ToString()},//ширина
                                    new shop_product_attribute(){ product_id = int.Parse(art), attribute_id = 13, text = height.ToString()},//высота
                                    new shop_product_attribute(){ product_id = int.Parse(art), attribute_id = 14, text = radius},//радиус
                                    new shop_product_attribute(){ product_id = int.Parse(art), attribute_id = 15, text = season},//сезон, у меня - зима или лето
                                    new shop_product_attribute(){ product_id = int.Parse(art), attribute_id = 16, text = inCol},//ИН
                                    new shop_product_attribute(){ product_id = int.Parse(art), attribute_id = 17, text = isCol},//ИС
                                    new shop_product_attribute(){ product_id = int.Parse(art), attribute_id = 18, text = rft},//РанФлэт
                                    new shop_product_attribute(){ product_id = int.Parse(art), attribute_id = 19, text = gruz},//Грузовой
                                    new shop_product_attribute(){ product_id = int.Parse(art), attribute_id = 20, text = spike},//Шипы
                                };
                                shop_product_to_category sptc = new shop_product_to_category() {
                                    category_id = 1,
                                    product_id = id
                                };
                                shop_product_to_layout sptl = new shop_product_to_layout() {
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
                                switch (season) {
                                    case "Лето":
                                        seasonToMetaTitle = "Летние";
                                        break;
                                    case "Зима":
                                        seasonToMetaTitle = "Зимние";
                                        break;
                                }
                                string productName = manufacturer + " " + model + " " + width + "/" + height + radius + " " + inCol + isCol + " " + rft + spike;
                                shop_product_description spd = new shop_product_description() {
                                    product_id = id,
                                    language_id = 1,
                                    name = productName,
                                    description = "",
                                    tag = width + "," + height + "," + radius,
                                    meta_title = seasonToMetaTitle + " шины " + productName + ". Магазин автошин TireShop",
                                    meta_description = seasonToMetaTitle+" шины "+productName + " по "+price*4+ " руб/шт. Доставка по СПб и в другие регионы",
                                    meta_keyword = ""
                                };
                                shop_url_alias sua = new shop_url_alias()
                                {
                                    query = "product_id="+id,
                                    keyword = manufacturer+"_"+id
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
                        else
                            MessageBox.Show("Такой товар уже есть в базе!", "Информация", MessageBoxButton.OK);
                    }
                    
                }
                catch (DbEntityValidationException ex) {
                    foreach (var eve in ex.EntityValidationErrors)
                    {
                        Console.WriteLine("Entity of type \"{0}\" in state \"{1}\" has the following validation errors:",
                            eve.Entry.Entity.GetType().Name, eve.Entry.State);
                        foreach (var ve in eve.ValidationErrors)
                        {
                            Console.WriteLine("- Property: \"{0}\", Error: \"{1}\"",
                                ve.PropertyName, ve.ErrorMessage);
                        }
                    }
                    throw;
                    log.Error(ex.Message + " \n" + ex.StackTrace);
                    MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
                }
            }
            else {
                MessageBox.Show("Не заполнено одно или несколько полей формы добавления продукта!", "Информация", MessageBoxButton.OK);
            }

        }

        private void CategoriesInModels_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var cmb = sender as ComboBox;
            int category = (int)cmb.SelectedValue;
            using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
            {
                var manufacturers = db.manufacturers.Where(w => w.CategoryId == category).ToList();
                Manufacturers.ItemsSource = manufacturers;
                Manufacturers.SelectedValuePath = "ManufacturerId";
                Manufacturers.DisplayMemberPath = "ManufacturerName";
            }
        }

        private void Categories2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var cmb = sender as ComboBox;
            int category = (int)cmb.SelectedValue;
            using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
            {
                var manufacturers = db.manufacturers.Where(w => w.CategoryId == category).ToList();
                Manufacturers2.ItemsSource = manufacturers;
                Manufacturers2.SelectedValuePath = "ManufacturerId";
                Manufacturers2.DisplayMemberPath = "ManufacturerName";
                Manufacturers2.IsEnabled = true;
                Season.ItemsSource = new List<Season>() { new Season() { Id = 1, Name = "Зима" }, new Season() { Id = 2, Name = "Лето" } };
                Season.DisplayMemberPath = "Name";
                Season.SelectedValuePath = "Id";
                Gruz.ItemsSource = new List<Gruz>() { new Gruz() { Id = 1, Name = "Да" }, new Gruz() { Id = 2, Name = "Нет" } };
                Gruz.DisplayMemberPath = "Name";
                Gruz.SelectedValuePath = "Id";
                RFT.ItemsSource = new List<RFT>() { new RFT() { Id = 1, Name = "Да" }, new RFT() { Id = 2, Name = "Нет" } };
                RFT.DisplayMemberPath = "Name";
                RFT.SelectedValuePath = "Id";
                Spikes.ItemsSource = new List<Spike>() { new Spike() { Id = 1, Name = "шип" }, new Spike() { Id = 2, Name = "Нет" } };
                Spikes.DisplayMemberPath = "Name";
                Spikes.SelectedValuePath = "Id";
                //Articul.Text = GenerateArticul().ToString();
            }
        }

        private void Manufacturers2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int category = (int)Categories2.SelectedValue;
            int ManId = (int)Manufacturers2.SelectedValue;
            using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
            {
                var models = db.models.Where(w => w.CategoryId == category && w.ManufacturerId==ManId).ToList();
                Models.SelectionChanged -= Models_SelectionChanged;
                Models.ItemsSource = models;
                Models.SelectedValuePath = "ModelId";
                Models.DisplayMemberPath = "ModelName";
                Models.IsEnabled = true;
                Models.SelectionChanged += Models_SelectionChanged;
            }
        }

        private string GenerateProdNumber(int cat, int manuf, int model) {
            using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities()) {
                string val = "";
                if (db.products.Any(a => a.CategoryId == cat && a.ManufacturerId == manuf && a.ModelId == model))
                {
                    var tmpList = db.products.Where(w => w.CategoryId == cat && w.ManufacturerId == manuf && w.ModelId == model).ToList();
                    int length = tmpList.Count;
                    var prod = tmpList[length - 1];
                    //int tmpid = int.Parse(prod.ProdNumber);
                    var tmp = prod.ProdNumber.Substring(prod.ProdNumber.Length - 1);
                    var tmpCount = int.Parse(tmp)+1;
                    char tmpC = Convert.ToChar(tmp);
                    val = prod.ProdNumber.Replace(tmp.ToString(), tmpCount.ToString());
                }
                else {
                    string manufStr = "", modelStr = "", count = "00001", ProdNumber = "";
                    if (manuf < 10)
                        manufStr = "0" + manuf;
                    else
                        manufStr = manuf.ToString();
                    if (model < 10)
                        modelStr = "0" + model;
                    else
                        modelStr = model.ToString();
                    val = cat.ToString() + manufStr + modelStr + count;
                }
                return val;
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            _mainWnd.mainFrame.Refresh();
        }

        private void SelectFile_Click(object sender, RoutedEventArgs e)
        {
            
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel file (*.xlsx)|*.xlsx";
            if (ofd.ShowDialog() == true) {
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities()) {
                    try {
                        backgroundWorker = new BackgroundWorker();
                        bgWork idol = new bgWork() { fileName = ofd.FileName };
                        backgroundWorker.WorkerReportsProgress = true;
                        backgroundWorker.DoWork += DoWork;
                        backgroundWorker.ProgressChanged += ProgressChanged;
                        backgroundWorker.RunWorkerCompleted += BackgroundWorker_RunWorkerCompleted;
                        backgroundWorker.RunWorkerAsync(idol);
                    }
                    catch (Exception ex)
                    {
                        int point2 = 0;
                    }
                }
            }
        }

        private void DoWork(object sender, DoWorkEventArgs e)
        {
            bgWork idol = (bgWork)e.Argument;
            List<XlsxImport> stations = GetListOfProducts(idol.fileName, backgroundWorker);
            e.Result = stations;
        }

        private void ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // This is called on the UI thread when ReportProgress method is called
            progressBar.Value += e.ProgressPercentage;
            if (progressBar.Value == 100) {
                progressBar.Value = 0;
            }
        }
        /// <summary>
        /// отвечает за лоадер
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                // Ошибка была сгенерирована обработчиком события DoWork
                MessageBox.Show(e.Error.Message, "Произошла ошибка");
            }
            else if (e.Cancelled)
            {
                int point = 0;
            }
            else//в случае нормальной работы основного метода все кнопки становятся доступны пользователю, изначально все залочены
            {
                List<XlsxImport> itemSource = (List<XlsxImport>)e.Result;
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    int counter = 1;
                    int tmpManufacturer = 0, tmpModel = 0;
                    foreach (var value in itemSource.OrderBy(o=>o.Manufacturer).ThenBy(t=>t.Model))
                    {

                        decimal optPrice = 0, price = 0, purPrice = 0;
                        if (value.OptPrice.HasValue)
                            optPrice = (decimal)value.OptPrice;
                        if (value.PurPrice != null)
                            purPrice = (decimal)value.PurPrice;
                        if (value.Price != null)
                            price = (decimal)value.Price;
                        if (!db.products.Any(a => a.ManufacturerId == value.Manufacturer && a.ModelId == value.Model && a.Radius == value.Radius && a.ProdNumber == value.Articul &&
                            a.Width == value.Width && a.Height == value.Height && a.InCol == value.InCol && a.IsCol == value.IsCol && a.Gruz == value.Gruz &&
                            a.OptPrice == optPrice && a.Price == price && a.PurchasePrice == purPrice && a.RFT == value.RFT && a.Season == value.Season && a.Spikes == value.Spikes))
                        {
                            string manuf = "", modelI = "", count = "", ProdNumber = "";
                            if (value.Manufacturer < 10)
                                manuf = "0" + value.Manufacturer;
                            else
                                manuf = value.Manufacturer.ToString();
                            if (value.Model < 10)
                                modelI = "0" + value.Model;
                            else
                                modelI = value.Model.ToString();
                            if (counter < 10)
                                count = "0000" + counter;
                            else if (counter < 100 && counter > 10)
                                count = "000" + counter;
                            else if (counter < 1000 && counter > 100)
                                count = "00" + counter;
                            else if (counter < 10000 && counter > 1000)
                                count = "0" + counter;
                            if (value.Manufacturer != tmpManufacturer && value.Model != tmpModel)
                                counter = 1;
                            else
                                counter++;
                            tmpManufacturer = value.Manufacturer;
                            tmpModel = value.Model;
                            ProdNumber = value.CategoryId + "" + manuf + modelI + count;
                            var prod = new product()
                            {
                                ProdNumber = ProdNumber,
                                CategoryId = 1,//потом перебьем, щас то одна категория всего
                                Gruz = value.Gruz,
                                Height = value.Height,
                                InCol = value.InCol,
                                IsCol = value.IsCol,
                                ManufacturerId = value.Manufacturer,
                                ModelId = value.Model,
                                OptPrice = optPrice,
                                Price = price,
                                PurchasePrice = purPrice,
                                Radius = value.Radius,
                                RFT = value.RFT,
                                Season = value.Season,
                                Spikes = value.Spikes,
                                Width = value.Width
                            };
                            db.products.Add(prod);
                            try
                            {
                                db.SaveChanges();

                            }
                            catch (DbEntityValidationException ex)
                            {
                                foreach (var eve in ex.EntityValidationErrors)
                                {
                                    Console.WriteLine("Entity of type \"{0}\" in state \"{1}\" has the following validation errors:",
                                        eve.Entry.Entity.GetType().Name, eve.Entry.State);
                                    foreach (var ve in eve.ValidationErrors)
                                    {
                                        Console.WriteLine("- Property: \"{0}\", Error: \"{1}\"",
                                            ve.PropertyName, ve.ErrorMessage);
                                    }
                                }
                                throw;
                            }
                            int totalQuant = 0;
                            foreach (var store in value.Storehouse)
                            {
                                var pq = new productquantity()
                                {
                                    ProductId = prod.ProductId,//store.ProductId,
                                    Quantity = store.Quantity,
                                    StorehouseId = store.StorehouseId
                                };
                                totalQuant += (int)store.Quantity;
                                if (totalQuant > 0)
                                    prod.ProdStatus = true;
                                else
                                    prod.ProdStatus = false;
                                db.Entry(prod).Property(p => p.ProdStatus).IsModified = true;
                                db.productquantities.Add(pq);
                                db.SaveChanges();
                            }
                            using (u0324292_mainEntities db2 = new u0324292_mainEntities())
                            {
                                string manufacturer = "", model = "";
                                manufacturer = db.manufacturers.Single(s => s.ManufacturerId == value.Manufacturer).ManufacturerName;
                                model = db.models.Single(s => s.ModelId == value.Model).ModelName;
                                int shopStat = 0;
                                if (totalQuant > 0)//больше нуля - продаем
                                    shopStat = 7;
                                else//в любом другом случае предзаказ
                                    shopStat = 8;
                                var id = uint.Parse(prod.ProdNumber);
                                shop_product sp = new shop_product()
                                {
                                    product_id = id,//айдишник продукта в сайтовой базе
                                    model = model,//наименование модели
                                    quantity = totalQuant,//количество в наличии
                                    stock_status_id = shopStat,//статус
                                    manufacturer_id = db2.shop_manufacturer.Single(s => s.name == manufacturer).manufacturer_id,//получение и присвоение айдишника производителя
                                    price = price,//цена
                                    status = true,
                                    sku = "",
                                    upc = "",
                                    ean = "",
                                    jan = "",
                                    isbn = "",
                                    mpn = "",
                                    location = ""
                                    //subtract = true,

                                };
                                List<shop_product_attribute> lst = new List<shop_product_attribute>()//список атрибутов
                                {
                                    new shop_product_attribute(){ product_id = int.Parse(prod.Articul), attribute_id = 12, text = prod.Width.ToString()},//ширина
                                    new shop_product_attribute(){ product_id = int.Parse(prod.Articul), attribute_id = 13, text = prod.Height.ToString()},//высота
                                    new shop_product_attribute(){ product_id = int.Parse(prod.Articul), attribute_id = 14, text = prod.Radius},//радиус
                                    new shop_product_attribute(){ product_id = int.Parse(prod.Articul), attribute_id = 15, text = prod.Season},//сезон, у меня - зима или лето
                                    new shop_product_attribute(){ product_id = int.Parse(prod.Articul), attribute_id = 16, text = prod.InCol},//ИН
                                    new shop_product_attribute(){ product_id = int.Parse(prod.Articul), attribute_id = 17, text = prod.IsCol},//ИС
                                    new shop_product_attribute(){ product_id = int.Parse(prod.Articul), attribute_id = 18, text = prod.RFT},//РанФлэт
                                    new shop_product_attribute(){ product_id = int.Parse(prod.Articul), attribute_id = 19, text = prod.Gruz},//Грузовой
                                    new shop_product_attribute(){ product_id = int.Parse(prod.Articul), attribute_id = 20, text = prod.Spikes},//Шипы
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
                                switch (prod.Season)
                                {
                                    case "Лето":
                                        seasonToMetaTitle = "Летние";
                                        break;
                                    case "Зима":
                                        seasonToMetaTitle = "Зимние";
                                        break;
                                }
                                string productName = manufacturer + " " + model + " " + prod.Width + "/" + prod.Height + prod.Radius + " " + prod.InCol + prod.IsCol + " " + prod.RFT + prod.Spikes;
                                shop_product_description spd = new shop_product_description()
                                {
                                    product_id = id,
                                    language_id = 1,
                                    name = productName,
                                    description = "",
                                    tag = prod.Width + "," + prod.Height + "," + prod.Radius,
                                    meta_title = seasonToMetaTitle + " шины " + productName + ". Магазин автошин TireShop",
                                    meta_description = seasonToMetaTitle + " шины " + productName + " по " + price * 4 + " руб/шт. Доставка по СПб и в другие регионы",
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
                    int stop = 0;
                    
                    MessageBox.Show("Данные из файла успешно импортированы в базу данных!", "Информация", MessageBoxButton.OK);
                }
                backgroundWorker.Dispose();
            }
        }

        private List<XlsxImport> GetListOfProducts(string fileName, BackgroundWorker worker) {
            List<XlsxImport> lst = new List<XlsxImport>();
            using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fileName, false))
                {
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    SharedStringTable sst = sstpart.SharedStringTable;

                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    Worksheet sheet = worksheetPart.Worksheet;

                    //var cells = sheet.Descendants<Cell>();
                    var rows = sheet.Descendants<Row>().ToList();
                    int eCount = rows.Count();
                    int iteration =  eCount/ 100;

                    for (int i = 1; i < rows.Count(); i++)//foreach (Row row in rows)
                    {
                        if (worker != null)//возврат значений для лоадера
                        {
                            if (worker.CancellationPending)
                            {
                                // Возврат без какой-либо дополнительной работы
                                return null;
                            }

                            if (worker.WorkerReportsProgress)
                            {
                                //float progress = ((float)(i + 1)) / list.Length * 100;
                                if (i != rows.Count() - 1)
                                    worker.ReportProgress(iteration);
                                else
                                    worker.ReportProgress(iteration + 2);
                                //(int)Math.Round(progress));
                            }
                        }
                        var row = rows[i];
                        var item = new XlsxImport() { CategoryId = 1};
                        item.Storehouse = new List<ProductQuantity>();
                        int tmpManId = -1;
                        var cells = row.Elements<Cell>().ToList();
                        for (int j = 0; j < cells.Count; j++)
                        {
                            var c = cells[j];
                            string text = "";
                            if (c.CellValue != null)
                                text = c.CellValue.Text;
                            else
                                text = "";
                            if (c.DataType != null && c.DataType == CellValues.SharedString)
                            {
                                int ssid = int.Parse(text);
                                string str = sst.ChildElements[ssid].InnerText;
                                if (str != string.Empty)
                                {
                                    if (j == 2)
                                    {
                                        var manExist = db.manufacturers.Where(w => w.ManufacturerName == str).ToList();
                                        if (manExist.Count > 0)
                                        {
                                            var manuf = db.manufacturers.Single(s => s.ManufacturerName == str);
                                            item.Manufacturer = manuf.ManufacturerId;
                                            tmpManId = manuf.ManufacturerId;
                                        }
                                        else
                                        {
                                            manufacturer man = new manufacturer()
                                            {
                                                ManufacturerName = str,
                                                CategoryId = 1
                                            };
                                            db.manufacturers.Add(man);
                                            db.SaveChanges();
                                            item.Manufacturer = man.ManufacturerId;
                                            tmpManId = man.ManufacturerId;
                                        }
                                    }
                                    if (j == 3)
                                    {
                                        var modelExist = db.models.Where(w => w.ModelName == str && w.ManufacturerId == tmpManId).ToList();
                                        if (modelExist.Count > 0)
                                        {
                                            var m = db.models.Single(s => s.ModelName == str && s.ManufacturerId == tmpManId);
                                            item.Model = m.ModelId;
                                        }
                                        else
                                        {
                                            model man = new model()
                                            {
                                                ModelName = str,
                                                ManufacturerId = tmpManId,
                                                CategoryId = 1
                                            };
                                            db.models.Add(man);
                                            db.SaveChanges();
                                            item.Model = man.ModelId;
                                        }
                                    }
                                    if (j == 6)
                                        item.Radius = str;
                                    if (j == 8)
                                        item.IsCol = str;
                                    if (j == 9)
                                        item.RFT = str;
                                    if (j == 10)
                                        item.Gruz = str;
                                    if (j == 11)
                                        item.Spikes = str;
                                    if (j == 12)
                                        item.Season = str;
                                }
                            }
                            else
                            {
                                if (text != string.Empty)
                                {
                                    if (j == 0)
                                        item.Articul = text;
                                    if (j == 4)
                                    {
                                        if (int.TryParse(text, out var tmp))
                                        {
                                            int width = int.Parse(text);
                                            item.Width = width;
                                        }
                                    }
                                    if (j == 5)
                                    {
                                        if (float.TryParse(text, out var tmp))
                                        {
                                            float height = float.Parse(text);
                                            item.Height = height;
                                        }
                                    }
                                    if (j == 7)
                                        item.InCol = text;
                                    if (j == 13)
                                    {
                                        bool res = int.TryParse(text, out var tmp);
                                        if (res)
                                        {
                                            var pq = new ProductQuantity()
                                            {
                                                //ProductId = number,
                                                Quantity = int.Parse(text),
                                                StorehouseName = "Гараж таллинское шоссе",
                                                StorehouseId = db.storehouses.Where(w => w.StorehouseName == "Гараж таллинское шоссе").Select(s => s.StorehouseId).First()
                                            };
                                            item.Storehouse.Add(pq);
                                        }
                                    }
                                    if (j == 14)
                                    {
                                        if (int.TryParse(text, out var tmp))
                                        {
                                            var pq = new ProductQuantity()
                                            {
                                                //ProductId = number,
                                                Quantity = int.Parse(text),
                                                StorehouseName = "Фрунзенская контейнер",
                                                StorehouseId = db.storehouses.Where(w => w.StorehouseName == "Фрунзенская контейнер").Select(s => s.StorehouseId).First()
                                            };
                                            item.Storehouse.Add(pq);
                                        }
                                    }
                                    if (j == 15)
                                    {
                                        if (int.TryParse(text, out var tmp))
                                        {
                                            item.Storehouse.Add(new ProductQuantity()
                                            {
                                                //ProductId = number,
                                                Quantity = int.Parse(text),
                                                StorehouseName = "Петроспирт Ангар Левый контик",
                                                StorehouseId = db.storehouses.Where(w => w.StorehouseName == "Петроспирт Ангар Левый контик").Select(s => s.StorehouseId).First()
                                            });
                                        }
                                    }
                                    if (j == 16)
                                    {
                                        if (int.TryParse(text, out var tmp))
                                        {
                                            item.Storehouse.Add(new ProductQuantity()
                                            {
                                                //ProductId = number,
                                                Quantity = int.Parse(text),
                                                StorehouseName = "Петроспирт Ангар Левый Бокс",
                                                StorehouseId = db.storehouses.Where(w => w.StorehouseName == "Петроспирт Ангар Левый Бокс").Select(s => s.StorehouseId).First()
                                            });
                                        }
                                    }
                                    if (j == 17)
                                    {
                                        if (int.TryParse(text, out var tmp))
                                        {
                                            item.Storehouse.Add(new ProductQuantity()
                                            {
                                                //ProductId = number,
                                                Quantity = int.Parse(text),
                                                StorehouseName = "Петроспирт Ангар Центральный контейнер",
                                                StorehouseId = db.storehouses.Where(w => w.StorehouseName == "Петроспирт Ангар Центральный контейнер").Select(s => s.StorehouseId).First()
                                            });
                                        }
                                    }
                                    if (j == 18)
                                    {
                                        if (int.TryParse(text, out var tmp))
                                        {
                                            item.Storehouse.Add(new ProductQuantity()
                                            {
                                                //ProductId = number,
                                                Quantity = int.Parse(text),
                                                StorehouseName = "Петроспирт Ангар Правый бокс",
                                                StorehouseId = db.storehouses.Where(w => w.StorehouseName == "Петроспирт Ангар Правый бокс").Select(s => s.StorehouseId).First()
                                            });
                                        }
                                    }
                                    if (j == 19)
                                    {
                                        if (int.TryParse(text, out var tmp))
                                        {
                                            item.Storehouse.Add(new ProductQuantity()
                                            {
                                                //ProductId = number,
                                                Quantity = int.Parse(text),
                                                StorehouseName = "Петроспирт Ангар Правый контейнер",
                                                StorehouseId = db.storehouses.Where(w => w.StorehouseName == "Петроспирт Ангар Правый контейнер").Select(s => s.StorehouseId).First()
                                            });
                                        }
                                    }
                                    if (j == 20)
                                    {
                                        if (int.TryParse(text, out var tmp))
                                        {
                                            item.Storehouse.Add(new ProductQuantity()
                                            {
                                                //ProductId = number,
                                                Quantity = int.Parse(text),
                                                StorehouseName = "Петроспирт Контейнер 1",
                                                StorehouseId = db.storehouses.Where(w => w.StorehouseName == "Петроспирт Контейнер 1").Select(s => s.StorehouseId).First()
                                            });
                                        }
                                    }
                                    if (j == 21)
                                    {
                                        bool res = int.TryParse(text, out var tmp);
                                        if (res)
                                        {
                                            item.Storehouse.Add(new ProductQuantity()
                                            {
                                                //ProductId = number,
                                                Quantity = int.Parse(text),
                                                StorehouseName = "Петроспирт Контейнер 2",
                                                StorehouseId = db.storehouses.Where(w => w.StorehouseName == "Петроспирт Контейнер 2").Select(s => s.StorehouseId).First()
                                            });
                                        }
                                    }
                                    if (j == 22)
                                    {
                                        if (int.TryParse(text, out var tmp))
                                        {
                                            item.Storehouse.Add(new ProductQuantity()
                                            {
                                                //ProductId = number,
                                                Quantity = int.Parse(text),
                                                StorehouseName = "Петроспирт Контейнер 3",
                                                StorehouseId = db.storehouses.Where(w => w.StorehouseName == "Петроспирт Контейнер 3").Select(s => s.StorehouseId).First()
                                            });
                                        }
                                    }
                                    if (j == 23)
                                    {
                                        if (int.TryParse(text, out var tmp))
                                        {
                                            item.Storehouse.Add(new ProductQuantity()
                                            {
                                                //ProductId = number,
                                                Quantity = int.Parse(text),
                                                StorehouseName = "Контейнер 1",
                                                StorehouseId = db.storehouses.Where(w => w.StorehouseName == "Контейнер 1").Select(s => s.StorehouseId).First()
                                            });
                                        }
                                    }
                                    if (j == 24)
                                    {
                                        if (int.TryParse(text, out var tmp))
                                        {
                                            item.Storehouse.Add(new ProductQuantity()
                                            {
                                                // ProductId = number,
                                                Quantity = int.Parse(text),
                                                StorehouseName = "Контейнер 2",
                                                StorehouseId = db.storehouses.Where(w => w.StorehouseName == "Контейнер 2").Select(s => s.StorehouseId).First()
                                            });
                                        }
                                    }
                                    if (j == 25)
                                    {
                                        if (int.TryParse(text, out var tmp))
                                        {
                                            item.Storehouse.Add(new ProductQuantity()
                                            {
                                                //ProductId = number,
                                                Quantity = int.Parse(text),
                                                StorehouseName = "Контейнер 3",
                                                StorehouseId = db.storehouses.Where(w => w.StorehouseName == "Контейнер 3").Select(s => s.StorehouseId).First()
                                            });
                                        }
                                    }
                                    if (j == 26)
                                    {
                                        if (int.TryParse(text, out var tmp))
                                        {
                                            item.Storehouse.Add(new ProductQuantity()
                                            {
                                                //ProductId = number,
                                                Quantity = int.Parse(text),
                                                StorehouseName = "Контейнер 4",
                                                StorehouseId = db.storehouses.Where(w => w.StorehouseName == "Контейнер 4").Select(s => s.StorehouseId).First()
                                            });
                                        }
                                    }
                                    if (j == 27)
                                    {
                                        if (int.TryParse(text, out var tmp))
                                        {
                                            item.Storehouse.Add(new ProductQuantity()
                                            {
                                                // ProductId = number,
                                                Quantity = int.Parse(text),
                                                StorehouseName = "Контейнер 5",
                                                StorehouseId = db.storehouses.Where(w => w.StorehouseName == "Контейнер 5").Select(s => s.StorehouseId).First()
                                            });
                                        }
                                    }
                                    if (j == 28)
                                    {
                                        if (int.TryParse(text, out var tmp))
                                        {
                                            item.Storehouse.Add(new ProductQuantity()
                                            {
                                                //ProductId = number,
                                                Quantity = int.Parse(text),
                                                StorehouseName = "Контейнер 6",
                                                StorehouseId = db.storehouses.Where(w => w.StorehouseName == "Контейнер 6").Select(s => s.StorehouseId).First()
                                            });
                                        }
                                    }
                                    if (j == 29)
                                    {
                                        if (int.TryParse(text, out var tmp))
                                        {
                                            item.Storehouse.Add(new ProductQuantity()
                                            {
                                                //ProductId = number,
                                                Quantity = int.Parse(text),
                                                StorehouseName = "Контейнер 7",
                                                StorehouseId = db.storehouses.Where(w => w.StorehouseName == "Контейнер 7").Select(s => s.StorehouseId).First()
                                            });
                                        }
                                    }
                                    if (j == 30)
                                    {
                                        if (int.TryParse(text, out var tmp))
                                        {
                                            item.Storehouse.Add(new ProductQuantity()
                                            {
                                                //ProductId = number,
                                                Quantity = int.Parse(text),
                                                StorehouseName = "Задний бокс",
                                                StorehouseId = db.storehouses.Where(w => w.StorehouseName == "Задний бокс").Select(s => s.StorehouseId).First()
                                            });
                                        }
                                    }
                                    if (j == 31)
                                    {
                                        if (int.TryParse(text, out var tmp))
                                        {
                                            item.Storehouse.Add(new ProductQuantity()
                                            {
                                                // ProductId = number,
                                                Quantity = int.Parse(text),
                                                StorehouseName = "Витрины",
                                                StorehouseId = db.storehouses.Where(w => w.StorehouseName == "Витрины").Select(s => s.StorehouseId).First()
                                            });
                                        }
                                    }
                                    if (j == 32)
                                    {
                                        if (int.TryParse(text, out var tmp))
                                        {
                                            item.Storehouse.Add(new ProductQuantity()
                                            {
                                                // ProductId = number,
                                                Quantity = int.Parse(text),
                                                StorehouseName = "Улица ",
                                                StorehouseId = db.storehouses.Where(w => w.StorehouseName == "Улица ").Select(s => s.StorehouseId).First()
                                            });
                                        }
                                    }
                                    if (j == 33)
                                    {
                                        if (int.TryParse(text, out var tmp))
                                        {
                                            item.Storehouse.Add(new ProductQuantity()
                                            {
                                                // ProductId = number,
                                                Quantity = int.Parse(text),
                                                StorehouseName = "Петроспирт Контейнер 4",
                                                StorehouseId = db.storehouses.Where(w => w.StorehouseName == "Петроспирт Контейнер 4").Select(s => s.StorehouseId).First()
                                            });
                                        }
                                    }
                                    if (j == 34)
                                    {
                                        if (int.TryParse(text, out var tmp))
                                        {
                                            item.Storehouse.Add(new ProductQuantity()
                                            {
                                                // ProductId = number,
                                                Quantity = int.Parse(text),
                                                StorehouseName = "Будка",
                                                StorehouseId = db.storehouses.Where(w => w.StorehouseName == "Будка").Select(s => s.StorehouseId).First()
                                            });
                                        }
                                    }
                                    if (j == 35)
                                    {
                                        bool res = decimal.TryParse(text, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out var tmp);
                                        if (res)
                                        {
                                            item.PurPrice = decimal.Parse(text, CultureInfo.InvariantCulture);
                                        }
                                    }
                                    if (j == 36)
                                    {
                                        bool res = decimal.TryParse(text, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out var tmp);
                                        if (res)
                                        {
                                            item.Price = decimal.Parse(text, CultureInfo.InvariantCulture);
                                        }
                                    }
                                    if (j == 37)
                                    {
                                        bool res = decimal.TryParse(text, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out var tmp);
                                        if (res)
                                        {
                                            item.OptPrice = decimal.Parse(text, CultureInfo.InvariantCulture);
                                        }
                                    }
                                }
                            }
                        }
                        lst.Add(item);
                    }    
                    int point = 0;
                }
            }
            return lst;
        }

        private void ModelsEdit_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var cmb = sender as ComboBox;
            try
            {
                int prodId = (int)cmb.SelectedValue;
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    SeasonEdit.ItemsSource = new List<Season>() { new Season() { Id = 1, Name = "Зима" }, new Season() { Id = 2, Name = "Лето" } };
                    SeasonEdit.DisplayMemberPath = "Name";
                    SeasonEdit.SelectedValuePath = "Id";
                    GruzEdit.ItemsSource = new List<Gruz>() { new Gruz() { Id = 1, Name = "Да" }, new Gruz() { Id = 2, Name = "Нет" } };
                    GruzEdit.DisplayMemberPath = "Name";
                    GruzEdit.SelectedValuePath = "Id";
                    RFTEdit.ItemsSource = new List<RFT>() { new RFT() { Id = 1, Name = "Да" }, new RFT() { Id = 2, Name = "Нет" } };
                    RFTEdit.DisplayMemberPath = "Name";
                    RFTEdit.SelectedValuePath = "Id";
                    SpikesEdit.ItemsSource = new List<Spike>() { new Spike() { Id = 1, Name = "шип" }, new Spike() { Id = 2, Name = "Нет" } };
                    SpikesEdit.DisplayMemberPath = "Name";
                    SpikesEdit.SelectedValuePath = "Id";
                    var prod = db.products.Single(s => s.ProductId == prodId);
                    WidthEdit.Text = prod.Width.ToString();
                    if (Math.Truncate(prod.Height) - prod.Height < 0)
                        HeightEdit.Text = prod.Height.ToString("0.0", CultureInfo.InvariantCulture);
                    else
                        HeightEdit.Text = prod.Height.ToString();
                    RadiusEdit.Text = prod.Radius;
                    InColEdit.Text = prod.InCol;
                    IsColEdit.Text = prod.IsCol;
                    CountryEdit.Text = prod.Country;
                    GruzEdit.Text = prod.Gruz;
                    SpikesEdit.Text = prod.Spikes;
                    RFTEdit.Text = prod.RFT;
                    SeasonEdit.Text = prod.Season;
                    PriceEdit.Text = prod.Price.ToString();
                    PurchasePriceEdit.Text = prod.PurchasePrice.ToString();
                    OptPriceEdit.Text = prod.OptPrice.ToString();

                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void EditProductBtn_Click(object sender, RoutedEventArgs e)
        {
            if (ModelsEdit.SelectedValue != null) {
                int prodId = (int)ModelsEdit.SelectedValue;
                bool flag = false;
                try
                {
                    using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                    {
                        var prod = db.products.Single(s => s.ProductId == prodId);
                        if (WidthEdit.Text != prod.Width.ToString())
                        {
                            prod.Width = int.Parse(WidthEdit.Text);
                            db.Entry(prod).Property(p => p.Width).IsModified = true;
                            flag = true;
                        }
                        if (HeightEdit.Text != prod.Height.ToString())
                        {
                            prod.Height = float.Parse(HeightEdit.Text, CultureInfo.InvariantCulture);
                            db.Entry(prod).Property(p => p.Height).IsModified = true;
                            flag = true;
                        }
                        if (RadiusEdit.Text != prod.Radius)
                        {
                            prod.Radius = RadiusEdit.Text;
                            db.Entry(prod).Property(p => p.Radius).IsModified = true;
                            flag = true;
                        }
                        if (InColEdit.Text != prod.InCol)
                        {
                            prod.InCol = InColEdit.Text;
                            db.Entry(prod).Property(p => p.InCol).IsModified = true;
                            flag = true;
                        }
                        if (IsColEdit.Text != prod.IsCol)
                        {
                            prod.IsCol = IsColEdit.Text;
                            db.Entry(prod).Property(p => p.IsCol).IsModified = true;
                            flag = true;
                        }
                        if (CountryEdit.Text != prod.Country)
                        {
                            prod.Country = CountryEdit.Text;
                            db.Entry(prod).Property(p => p.Country).IsModified = true;
                            flag = true;
                        }
                        if (GruzEdit.Text != prod.Gruz)
                        {
                            prod.Gruz = GruzEdit.Text;
                            db.Entry(prod).Property(p => p.Gruz).IsModified = true;
                            flag = true;
                        }
                        if (SpikesEdit.Text != prod.Spikes)
                        {
                            prod.Spikes = SpikesEdit.Text;
                            db.Entry(prod).Property(p => p.Spikes).IsModified = true;
                            flag = true;
                        }
                        if (RFTEdit.Text != prod.RFT)
                        {
                            prod.RFT = RFTEdit.Text;
                            db.Entry(prod).Property(p => p.RFT).IsModified = true;
                            flag = true;
                        }
                        if (SeasonEdit.Text != prod.Season)
                        {
                            prod.Season = SeasonEdit.Text;
                            db.Entry(prod).Property(p => p.Season).IsModified = true;
                            flag = true;
                        }
                        if (PriceEdit.Text != prod.Price.ToString())
                        {
                            var newPrice = decimal.Parse(PriceEdit.Text, CultureInfo.InvariantCulture);
                            prod.Price = newPrice;
                            db.Entry(prod).Property(p => p.Price).IsModified = true;
                            //using (u0324292_mainEntities db2 = new u0324292_mainEntities()) {
                            //    var id = int.Parse(prod.ProdNumber);
                            //    var shopProd = db2.shop_product.Single(s => s.product_id == id);
                            //    shopProd.price = newPrice;
                            //    db2.Entry(shopProd).Property(p => p.price).IsModified = true;
                            //    db2.SaveChanges();
                            //}
                            flag = true;
                        }
                        if (PurchasePriceEdit.Text != prod.PurchasePrice.ToString())
                        {
                            prod.PurchasePrice = decimal.Parse(PurchasePriceEdit.Text, CultureInfo.InvariantCulture);
                            db.Entry(prod).Property(p => p.PurchasePrice).IsModified = true;
                            flag = true;
                        }
                        if (OptPriceEdit.Text != prod.OptPrice.ToString())
                        {
                            prod.OptPrice = decimal.Parse(OptPriceEdit.Text, CultureInfo.InvariantCulture);
                            db.Entry(prod).Property(p => p.OptPrice).IsModified = true;
                            flag = true;
                        }
                        if (flag)
                        {
                            try
                            {
                                db.SaveChanges();
                            }
                            catch (Exception ex)
                            {
                                log.Error(ex.Message + " \n" + ex.StackTrace);
                                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
                            }
                            WidthEdit.Text = "";
                            HeightEdit.Text = "";
                            RadiusEdit.Text = "";
                            InColEdit.Text = "";
                            IsColEdit.Text = "";
                            CountryEdit.Text = "";
                            GruzEdit.Text = "";
                            SpikesEdit.Text = "";
                            RFTEdit.Text = "";
                            SeasonEdit.Text = "";
                            PriceEdit.Text = "";
                            PurchasePriceEdit.Text = "";
                            OptPriceEdit.Text = "";
                            ModelsEdit.SelectionChanged -= ModelsEdit_SelectionChanged;
                            ModelsEdit.SelectedValue = -1;
                            ModelsEdit.SelectionChanged += ModelsEdit_SelectionChanged;
                            MessageBox.Show("Товар изменен!", "Информация", MessageBoxButton.OK);
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

        private void DeleteProductBtn_Click(object sender, RoutedEventArgs e)
        {
            if (ModelsEdit.SelectedValue != null)
            {
                int prodId = (int)ModelsEdit.SelectedValue;
                var res = MessageBox.Show("Вы действительно хотите полностью удалить этот товар? Действие необратимо.", "Информация", MessageBoxButton.OKCancel);
                if (res == MessageBoxResult.OK) {
                    try
                    {
                        bool flag = true;
                        using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                        {
                            var prod = db.products.Single(s => s.ProductId == prodId);
                            var operations = db.operations.ToList();
                            //db.products.Remove(prod);
                            prod.ProdStatus = false;
                            db.Entry(prod).Property(p => p.ProdStatus).IsModified = true;
                            db.SaveChanges();
                            WidthEdit.Text = "";
                            HeightEdit.Text = "";
                            RadiusEdit.Text = "";
                            InColEdit.Text = "";
                            IsColEdit.Text = "";
                            CountryEdit.Text = "";
                            GruzEdit.Text = "";
                            SpikesEdit.Text = "";
                            RFTEdit.Text = "";
                            SeasonEdit.Text = "";
                            PriceEdit.Text = "";
                            PurchasePriceEdit.Text = "";
                            OptPriceEdit.Text = "";
                            ModelsEdit.SelectionChanged -= ModelsEdit_SelectionChanged;
                            ModelsEdit.SelectedValue = -1;
                            ModelsEdit.SelectionChanged += ModelsEdit_SelectionChanged;
                            MessageBox.Show("Товар удален!", "Информация", MessageBoxButton.OK);
                            //using (u0324292_mainEntities db2 = new u0324292_mainEntities())
                            //{
                            //    var id = int.Parse(prod.ProdNumber);
                            //    if (db2.shop_product.Any(a => a.product_id == id))
                            //    {
                            //        var siteProd = db2.shop_product.Single(a => a.product_id == id);
                            //        //siteProd.quantity -= quant;
                            //        //db2.Entry(siteProd).Property(p => p.quantity).IsModified = true;
                            //        //if (siteProd.quantity >= 0)
                            //        //    db2.SaveChanges();
                            //        //else
                            //        //    log.Error(siteProd.model + siteProd.quantity + siteProd.product_id + " количество не может быть отрицательным");
                            //        //if (siteProd.quantity == 0 && siteProd.stock_status_id == 7)
                            //        //{
                            //            siteProd.status = false;
                            //            db2.Entry(siteProd).Property(p => p.stock_status_id).IsModified = true;
                            //            db2.SaveChangesAsync();
                            //        //}

                            //    }
                            //}
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

        private void ModelsEdit_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            var Cmb = sender as ComboBox;
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
                        return true;
                    }
                    else return false;
                }
            });
            itemsViewOriginal.Refresh();
        }

        private void Models_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            
            var cmb = sender as ComboBox;
            try
            {
                int category = (int)Categories2.SelectedValue;
                int ManId = (int)Manufacturers2.SelectedValue;
                int modelId = (int)cmb.SelectedValue;
                Articul.Text = GenerateProdNumber(category, ManId, modelId);
            }
            catch (Exception ex) {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }
    }
}
