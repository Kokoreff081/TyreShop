using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
using Spire.Xls;
using Microsoft.Win32;
using System.IO;
using System.Drawing.Printing;
using System.Printing;
using System.Drawing;
using NLog;
using System.Threading;

namespace Tyreshop
{
    /// <summary>
    /// Логика взаимодействия для SaleFixing.xaml
    /// </summary>
    public partial class SaleFixing : Window
    {

        private Logger log;
        private MainWindow _mainWnd;
        public List<PComboBox> list;
        private List<DGSaleItems> sales;
        private List<DGSaleItems> salesToCheck;
        private bool Savedoperation = false;
        public SaleFixing(MainWindow main, List<PComboBox> lpc, List<BdProducts>bdp)
        {
            InitializeComponent();
            log = LogManager.GetCurrentClassLogger();
            list = lpc;
            _mainWnd = main;
            sales = new List<DGSaleItems>();
            salesToCheck = new List<DGSaleItems>();
            Products.ItemsSource = list;
            Products.SelectedValuePath = "ProductId";
            Products.DisplayMemberPath = "ProductName";
            try
            {
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    var services = db.services.ToList();
                    Services.ItemsSource = services;
                    Services.SelectedValuePath = "ServiceId";
                    Services.DisplayMemberPath = "ServiceName";
                    var users = db.users.Where(w => w.Role == "manager" && w.UserId!=14).ToList();
                    Manager.ItemsSource = users;
                    Manager.SelectedValuePath = "UserId";
                    Manager.DisplayMemberPath = "UserName";
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }

        }

        private void AddToSaleBtn_Click(object sender, RoutedEventArgs e)
        {
            DelSaleButtonTag tag = new DelSaleButtonTag();
            var newSale = new DGSaleItems();
            var newSale2 = new DGSaleItems();
            int cardToTotalSum = -1;
            bool flag = false;
            if (CardPay.IsChecked == true)
            {
                var res = MessageBox.Show("Платеж на карте. Добавлять данную сумму в общую дневную выручку наличных?", "Информация", MessageBoxButton.OKCancel);
                if (res == MessageBoxResult.OK)
                    cardToTotalSum = 1;
                else
                    cardToTotalSum = 0;
            }
            else
                cardToTotalSum = 1;
            int managerId = 14;
            if (Manager.SelectedValue != null)
                managerId = (int)Manager.SelectedValue;
            try
            {
                
                if (Products.SelectedValue != null)
                {
                    tag.ServId = null;
                    tag.ProdId = (int)Products.SelectedValue;
                    tag.AnotherId = null;
                    tag.SaleNumber = sales.Count + 1;
                    newSale.ProductId = (int)Products.SelectedValue;
                    newSale.ProdName = Products.Text;
                    newSale.ProductId = (int)Products.SelectedValue;
                    newSale.Date = DateTime.Now.ToString("dd-MM-yyyy");
                    newSale.Time = DateTime.Now.ToString("hh:mm:ss");
                    newSale.Price = decimal.Parse(ProductPrice.Text);
                    newSale.Quantity = int.Parse(Quantity.Text);
                    newSale.SaleNumber = sales.Count + 1;
                    newSale.OperationType = "Продажа";
                    newSale.StoreId = (int)StorehouseFrom.SelectedValue;
                    newSale.CardToTotalSum = cardToTotalSum;
                    newSale.Comment = OperComment.Text;
                    newSale.UserId = managerId;
                    newSale.TagToBtn = tag;
                    if (PayTypeCheckBox.IsChecked == true)
                        newSale.PayType = "Безналичный расчет";
                    else
                        newSale.PayType = "Наличный расчет";
                    if (CardPay.IsChecked == true)
                        newSale.CardPayed = "Да";
                    else
                        newSale.CardPayed = "Нет";
                    if (!sales.Any(a => a.ProductId == (int)Products.SelectedValue))
                        sales.Add(newSale);
                    else
                    {
                        var sale = sales.Single(w => w.ProductId == (int)Products.SelectedValue);
                        sale.Quantity += int.Parse(Quantity.Text);
                        sale.Price += decimal.Parse(ProductPrice.Text);
                    }
                    DelSaleButtonTag tag2 = new DelSaleButtonTag()
                    {
                        ServId = null,
                        ProdId = (int)Products.SelectedValue,
                        AnotherId = null,
                        SaleNumber = sales.Count + 1
                    };
                    newSale2.TagToBtn = tag2;
                    newSale2.ProdName = Products.Text;
                    newSale2.ProductId = (int)Products.SelectedValue;
                    newSale2.Date = DateTime.Now.ToString("dd-MM-yyyy");
                    newSale2.Time = DateTime.Now.ToString("hh:mm:ss");
                    newSale2.Price = decimal.Parse(ProductPrice.Text);
                    newSale2.Quantity = int.Parse(Quantity.Text);
                    newSale2.SaleNumber = salesToCheck.Count + 1;
                    newSale2.OperationType = "Продажа";
                    newSale2.StoreId = (int)StorehouseFrom.SelectedValue;
                    newSale2.CardToTotalSum = cardToTotalSum;
                    newSale2.Comment = OperComment.Text;
                    newSale2.ProductId = (int)Products.SelectedValue;
                    newSale2.UserId = managerId;
                    if (PayTypeCheckBox.IsChecked == true)
                        newSale2.PayType = "Безналичный расчет";
                    else
                        newSale2.PayType = "Наличный расчет";
                    if (CardPay.IsChecked == true)
                        newSale2.CardPayed = "Да";
                    else
                        newSale2.CardPayed = "Нет";
                    salesToCheck.Add(newSale2);
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
            try
            {
                if (Services.SelectedValue != null)
                {
                    string servComment = "";
                    tag.ServId = (int)Services.SelectedValue;
                    tag.ProdId = null;
                    tag.AnotherId = null;
                    tag.SaleNumber = sales.Count + 1;
                    if (CardPayServ.IsChecked == true)
                    {
                        var res = MessageBox.Show("Платеж на карте. Добавлять данную сумму в общую дневную выручку наличных?", "Информация", MessageBoxButton.OKCancel);
                        if (res == MessageBoxResult.OK)
                            cardToTotalSum = 1;
                        else
                            cardToTotalSum = 0;
                    }
                    else
                        cardToTotalSum = 1;
                    int? servQuant = null;
                    bool res2 = int.TryParse(QuantityServ.Text, out var tmp);
                    bool Flag = false;
                    if (res2)
                    {
                        servQuant = int.Parse(QuantityServ.Text);
                        flag = true;
                    }
                    if (ServComment.Text != string.Empty)
                        servComment = ServComment.Text;
                    string PayType = "";
                    if (PayTypeCheckBoxServ.IsChecked == true)
                        PayType = "Безналичный расчет";
                    else
                        PayType = "Наличный расчет";
                    string CardPayed = "";
                    if (CardPayServ.IsChecked == true)
                        CardPayed = "Да";
                    else
                        CardPayed = "Нет";
                    if (flag)
                    {
                        newSale = new DGSaleItems()
                        {
                            ServiceId = (int)Services.SelectedValue,
                            Date = DateTime.Now.ToString("dd-MM-yyyy"),
                            Time = DateTime.Now.ToString("hh:mm:ss"),
                            ProdName = Services.Text,
                            Price = decimal.Parse(ServicePrice.Text),
                            SaleNumber = sales.Count + 1,
                            Quantity = (int)servQuant,
                            OperationType = "Услуга",
                            PayType = PayType,
                            StoreId = null,
                            CardPayed = CardPayed,
                            CardToTotalSum = cardToTotalSum,
                            TagToBtn = tag,
                            UserId = managerId,
                            Comment = servComment
                    };
                        sales.Add(newSale);
                        salesToCheck.Add(newSale);
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
            try
            {
                if (OtherProduct.Text != string.Empty)
                {

                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
            decimal totalSum = 0;
            if (sales.Any(a=>a.ProdName== "Общая сумма чека")) {
                var sale = sales.Single(s => s.ProdName == "Общая сумма чека");
                sales.RemoveAt(sales.IndexOf(sale));
            }
            foreach (var sale in sales) {
                totalSum += sale.Price;
            }
            sales.Add(new DGSaleItems() { ProdName = "Общая сумма чека", Price = totalSum });
            DGSaleFix.ItemsSource = sales;
            DGSaleFix.Items.Refresh();

            Products.SelectionChanged -= Products_SelectionChanged;
            Products.SelectedValue = -1;
            Products.SelectionChanged += Products_SelectionChanged;
            Products.Text = "";
            Quantity.TextChanged -= Quantity_TextChanged;
            Quantity.Text = string.Empty;
            Quantity.TextChanged += Quantity_TextChanged;
            ProductPrice.Text = string.Empty;
            StorehouseFrom.SelectedValue = -1;
            Services.SelectedValue = -1;
            Services.Text = "";
            ServicePrice.Text = "";
            SaleSave.IsEnabled = true;
            PrintSale.IsEnabled = true;
            Thread.Sleep(500);
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            ServiceSP.Visibility = Visibility.Visible;
            Adding.Height = new GridLength(Adding.ActualHeight + 100D, GridUnitType.Pixel); ;
        }

        private void CheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            ServiceSP.Visibility = Visibility.Collapsed;
            Adding.Height = new GridLength(Adding.ActualHeight - 100D, GridUnitType.Pixel); 
        }

        private void Products_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var cmb = sender as ComboBox;
            try
            {
                int prodId = (int)cmb.SelectedValue;
                decimal prodPrice = 0, optProdPrice=0;
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    prodPrice = (decimal)db.products.Where(w => w.ProductId == prodId).Select(s => s.Price).FirstOrDefault();
                    optProdPrice = (decimal)db.products.Where(w => w.ProductId == prodId).Select(s => s.OptPrice).FirstOrDefault();
                    var storehouses = db.productquantities.Where(w => w.ProductId == prodId && w.Quantity > 0).Join(db.storehouses, pq => pq.StorehouseId, s => s.StorehouseId, (pq, s) => new { StoreHouseId = pq.StorehouseId, StoreHouseName = s.StorehouseName + "(" + pq.Quantity + ")" }).ToList();
                    StorehouseFrom.ItemsSource = storehouses;
                    StorehouseFrom.SelectedValuePath = "StoreHouseId";
                    StorehouseFrom.DisplayMemberPath = "StoreHouseName";
                }
                ProductPrice.Text = prodPrice.ToString();
                OptPriceLbl.Content = optProdPrice.ToString();
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void Quantity_TextChanged(object sender, TextChangedEventArgs e)
        {
            var txt = sender as TextBox;
            try
            {
                if (decimal.TryParse(ProductPrice.Text, out var number))
                {
                    var price = decimal.Parse(ProductPrice.Text);
                    int newQuant = int.Parse(txt.Text);
                    var newPrice = (decimal)(price * newQuant);
                    ProductPrice.Text = newPrice.ToString();
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void NumberValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9]+");
            if(e.Text!=string.Empty)
                e.Handled = regex.IsMatch(e.Text);
        }

        private void SaveSale() {
            CultureInfo provider = CultureInfo.InvariantCulture;
            try
            {
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    var sNum = db.operations.ToList();
                    long number = 0;
                    if (sNum.Count == 0)
                        number = 1;
                    else
                        number = (long)sNum[sNum.Count - 1].SaleNumber + 1;
                    //int storeId = (int)StorehouseFrom.SelectedValue;
                    int quant = 0;
                    foreach (var item in salesToCheck)
                    {
                        string storehouseName = "";
                        if (item.StoreId != null)
                        {
                            var store = db.storehouses.Single(s => s.StorehouseId == item.StoreId);
                            storehouseName = store.StorehouseName;
                        }
                        operation oper = new operation()
                        {
                            OperationDate = DateTime.ParseExact(item.Date, "dd-MM-yyyy", provider),
                            OperationTime = DateTime.ParseExact(item.Time, "hh:mm:ss", provider),
                            Price = item.Price,
                            Count = item.Quantity,
                            ProductId = item.ProductId,
                            ServiceId = item.ServiceId,
                            SaleNumber = number,
                            PayType = item.PayType,
                            OperationType = item.OperationType,
                            Comment = item.Comment,
                            Storehouse = storehouseName,
                            CardPay = item.CardPayed,
                            CardToTotalSum = item.CardToTotalSum,
                            UserId = item.UserId
                        };
                        db.operations.Add(oper);
                        if (item.StoreId != null)
                        {
                            var store = db.productquantities.Single(s => s.ProductId == item.ProductId && s.StorehouseId == item.StoreId);
                            store.Quantity -= item.Quantity;
                            db.Entry(store).Property(p => p.Quantity).IsModified = true;
                            
                        }
                        if (item.ProductId != null && item.ProductId != 0) {
                            using (u0324292_mainEntities db2 = new u0324292_mainEntities())
                            {
                                var prod = db.products.Single(s => s.ProductId == item.ProductId);
                                var id = int.Parse(prod.ProdNumber);
                                if (db2.shop_product.Any(a => a.product_id == id))
                                {
                                    var siteProd = db2.shop_product.Single(a => a.product_id == id);
                                    siteProd.quantity -= item.Quantity;
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
                            var prodsQ = db.productquantities.Where(w => w.ProductId == item.ProductId).ToList();
                            foreach (var innerItem in prodsQ) {
                                quant += (int)innerItem.Quantity;
                            }
                            var product = db.products.Single(s => s.ProductId == item.ProductId);
                            if (quant > 0)
                                product.ProdStatus = true;
                            else
                                product.ProdStatus = false;
                            db.Entry(product).Property(p => p.ProdStatus).IsModified = true;
                        }
                    }
                    db.SaveChanges();
                    sales.Clear();
                    salesToCheck.Clear();
                    DGSaleFix.Items.Refresh();
                    SaleSave.IsEnabled = false;
                    PrintSale.IsEnabled = false;
                    Savedoperation = true;
                }

            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void PrintSale_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.Filter = "Excel Document(*.xlsx)|*.xlsx|Excel 2007-2010(*.xlsx)|*.xlsx";
            //SaveSale();
            string saleNumber = "";
            string saleDate = "";
            string productPrice = "";
            try
            {
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    int count = salesToCheck.Count();
                    var saleNumbers = db.operations.ToList();
                    if (Savedoperation)
                    {
                        var operation = saleNumbers[saleNumbers.Count - 1];
                        saleNumber = operation.SaleNumber.ToString();
                        saleDate = operation.OperationDate.ToString("m", CultureInfo.CurrentUICulture) + " " + operation.OperationDate.ToString("yyyy");
                        productPrice = db.products.Where(w => w.ProductId == operation.ProductId).Select(s => s.Price).FirstOrDefault().ToString();
                    }
                    else
                    {
                        var operation = salesToCheck[count - 1];
                        saleNumber = operation.SaleNumber.ToString();
                        saleDate = DateTime.Now.ToString("m", CultureInfo.CurrentUICulture) + " " + DateTime.Now.ToString("yyyy");//operation.OperationDate.ToString("m", CultureInfo.CurrentUICulture) + " " + operation.OperationDate.ToString("yyyy");
                        productPrice = db.products.Where(w => w.ProductId == operation.ProductId).Select(s => s.Price).FirstOrDefault().ToString();
                    }
                    int oint = 0;
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
            string header = @"Продавец: ИП Разуменко Анна Игоревна
ИНН: 390103431861
Расч.счет: 40802810500000463519
Банк: АО «Тинькофф Банк» Москва, 123060, 1-й Волоколамский проезд, д. 10, стр. 1
БИК:    044525974     Корр.счет:  30101810145250000974 
Адрес продавца: Санкт - Петербург ул.Калинина 5 " + "\"TireShop - новые шины по самым низким ценам\"" +
"\n Тел.8 - 965 - 818 - 84 - 46 сайт: EUROKOLESO.RU";
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            sheet.SetColumnWidth(1, 2);//a
            sheet.SetColumnWidth(2, 0.5);//b
            sheet.SetColumnWidth(3, 7.5);//c
            sheet.SetColumnWidth(4, 21.3);//d
            sheet.SetColumnWidth(5, 8.1);//e
            sheet.SetColumnWidth(6, 7.3);//f
            sheet.SetColumnWidth(7, 1.6);//g
            sheet.SetColumnWidth(8, 2.8);//h
            sheet.SetColumnWidth(9, 5.3);//i
            sheet.SetColumnWidth(10, 5.6);//j
            sheet.SetColumnWidth(11, 2);//k
            sheet.SetColumnWidth(12, 1.3);//l
            sheet.SetColumnWidth(13, 5.6);//m
            sheet.SetColumnWidth(14, 2.9);//n
            sheet.SetColumnWidth(15, 2.6);//o
            sheet.SetColumnWidth(16, 0.7);//p
            sheet.SetColumnWidth(17, 1);//r
            sheet.Range["A1:R8"].Merge();
            sheet.Range["A1"].Text = header;
            sheet.Range["C9:G9"].Merge();
            sheet.Range["C9"].Text = "Товарный чек № "+saleNumber + " от ";
            sheet.Range["H9:M9"].Merge();
            sheet.Range["H9"].Text = saleDate;
            sheet.Range["A11:B11"].Merge();
            sheet.Range["A11"].Text = "№";
            sheet.Range["A11:M11"].BorderInside();
            sheet.Range["A11:M11"].BorderAround();
            sheet.Range["A11"].Style.Color = System.Drawing.Color.LightGray;
            sheet.Range["C11:F11"].Merge();
            sheet.Range["C11"].Text = "Наименование товара";
            //sheet.Range["C11"].BorderInside();
            sheet.Range["C11"].Style.Color = System.Drawing.Color.LightGray;
            sheet.Range["G11:I11"].Merge();
            sheet.Range["G11:I11"].Style.Color = System.Drawing.Color.LightGray;
            sheet.Range["G11"].Text = "Кол-во";
            //sheet.Range["I11"].BorderInside();
            sheet.Range["I11"].Style.Color = System.Drawing.Color.LightGray;
            sheet.Range["J11:L11"].Merge();
            sheet.Range["J11"].Text = "Цена";
            //sheet.Range["J11"].BorderInside();
            sheet.Range["J11"].Style.Color = System.Drawing.Color.LightGray;
            sheet.Range["M11:R11"].Merge();
            sheet.Range["M11"].Text = "Сумма";
            sheet.Range["M11:R11"].BorderAround();
            sheet.Range["M11"].Style.Color = System.Drawing.Color.LightGray;

            int counter = 12;
            decimal totalSum = 0;
            for (int i = 0; i < sales.Count-1; i++) {
                var item = sales[i];
                sheet.Range["A"+counter+":B"+counter].Merge();
                sheet.Range["A" + counter].Text = item.SaleNumber.ToString();
                sheet.Range["C" + counter + ":F" + counter].Merge();
                sheet.Range["C" + counter].Text = item.ProdName;
                sheet.Range["G" + counter + ":I" + counter].Merge();
                sheet.Range["G" + counter].Text = item.Quantity.ToString();
                sheet.Range["J" + counter + ":L" + counter].Merge();
                if (item.ServiceId == 0)
                    wb.Worksheets[0].Range["J" + counter].Text = productPrice;
                else
                {
                    //var totalPrice = item.Price * item.Quantity;
                    sheet.Range["J" + counter].Text = ((decimal)(item.Price / item.Quantity)).ToString();
                }
                sheet.Range["M" + counter + ":R" + counter].Merge();
                sheet.Range["M" + counter].Text = item.Price.ToString() +" руб.";
                sheet.Range["A" + counter + ":R" + counter].BorderInside();
                sheet.Range["A" + counter + ":R" + counter].BorderAround();
                counter++;
                totalSum += item.Price;
            }
            sheet.Range["A" + counter + ":H" + counter].Merge();
            sheet.Range["I" + counter + ":L" + counter].Merge();
            sheet.Range["I" + counter].Text = "сумма чека: ";
            sheet.Range["M" + counter + ":R" + counter].Merge();
            sheet.Range["M" + counter].Text = totalSum.ToString() + " руб.";
            sheet.Range["M" + counter + ":R" + counter].BorderAround();
            counter++;
            sheet.Range["B" + counter + ":P" + counter].Merge();
            sheet.Range["B" + counter].Text = "Всего наименований " + sales.Count;
            counter += 3;
            sheet.Range["C" + counter + ":K" + (counter+4)].Merge();
            sheet.Range["C" + counter].Text = "Отпустил продавец:__________________________________";
            counter += 4;
            sheet.Range["D" + counter].Text = "Товар получен, осмотрен, претензий по внешнему виду не имею. О условиях гарантии предупрежден.";
            counter++;
            sheet.Range["F" + counter].Text = "М.П.";

            //bool? result = saveFile.ShowDialog();
            //if (result.HasValue && result.Value)
            //{
            //    using (Stream stream = saveFile.OpenFile())
            //    {
            //        wb.SaveToStream(stream);
            //    }

            //}
            //int counter2 = 0;
            //do
            //{
            //    System.Threading.Thread.Sleep(500);
            //    counter2++;
            //} while (!File.Exists(saveFile.FileName) && counter2 < 10);
            //if (counter2 < 10)
            //    System.Diagnostics.Process.Start(saveFile.FileName);
            PrintDialog dialog = new PrintDialog();
            dialog.UserPageRangeEnabled = true;
            PageRange rang = new PageRange(1, 1);
            dialog.PageRange = rang;
            PageRangeSelection seletion = PageRangeSelection.UserPages;
            dialog.PageRangeSelection = seletion;
            //dialog.PrintQueue = new PrintQueue(new PrintServer(), "Brother DCP-9020CDW Printer (копия 1)");
            PrintDocument pd = wb.PrintDocument;
            if (dialog.ShowDialog() == true)
            {
                pd.Print();
                //Thread.Sleep(10000);
                //int point = 0;
            }
            //this.Close();
        }

        private void QuantityServ_TextChanged(object sender, TextChangedEventArgs e)
        {
            var txt = sender as TextBox;
            try
            {
                if (decimal.TryParse(ServicePrice.Text, out var number))
                {
                    var price = decimal.Parse(ServicePrice.Text);
                    int newQuant = int.Parse(txt.Text);
                    var newPrice = (decimal)(price * newQuant);
                    ServicePrice.Text = newPrice.ToString();
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message + " \n" + ex.StackTrace);
                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
            }
        }

        private void SaleSave_Click(object sender, RoutedEventArgs e)
        {
            SaveSale();
            MessageBox.Show("Продажа успешно оформлена!", "Информация", MessageBoxButton.OK);
            //this.Close();
        }

        private void Products_KeyUp(object sender, KeyEventArgs e)
        {
            var Cmb = sender as ComboBox;
            CollectionView itemsViewOriginal = (CollectionView)CollectionViewSource.GetDefaultView(Cmb.ItemsSource);
            itemsViewOriginal.Filter = ((o) =>
            {
                if (String.IsNullOrEmpty(Cmb.Text)) return true;
                else
                {
                    var obj = o as PComboBox;
                    string filter_param = Cmb.Text;
                    if ((obj.ProductName).Contains(Cmb.Text))
                    {
                        Cmb.IsDropDownOpen = true;
                        if (Key.Enter == e.Key) {
                            int prodId = (int)Cmb.SelectedValue;
                            decimal prodPrice = 0;
                            try
                            {
                                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                                {
                                    prodPrice = (decimal)db.products.Where(w => w.ProductId == prodId).Select(s => s.Price).FirstOrDefault();
                                    var storehouses = db.productquantities.Where(w => w.ProductId == prodId && w.Quantity > 0).Join(db.storehouses, pq => pq.StorehouseId, s => s.StorehouseId, (pq, s) => new { StoreHouseId = pq.StorehouseId, StoreHouseName = s.StorehouseName }).ToList();
                                    StorehouseFrom.ItemsSource = storehouses;
                                    StorehouseFrom.SelectedValuePath = "StoreHouseId";
                                    StorehouseFrom.DisplayMemberPath = "StoreHouseName";
                                }
                            }
                            catch (Exception ex)
                            {
                                log.Error(ex.Message + " \n" + ex.StackTrace);
                                MessageBox.Show("Кажется, что-то пошло не так...", "Информация", MessageBoxButton.OK);
                            }
                            ProductPrice.Text = prodPrice.ToString();
                        }
                        //TextBox TxtBox = (TextBox)Cmb.Template.FindName("PART_EditableTextBox", Cmb);
                        //TxtBox.SelectionLength = filter_param.Length;
                        return true;
                    }
                    else return false;
                }
            });
            itemsViewOriginal.Refresh();
        }

        

        private void QuantityOtherProduct_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void DelSale_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            DelSaleButtonTag dt = btn.Tag as DelSaleButtonTag;
            try
            {
                DGSaleItems sale = new DGSaleItems();
                var lst = sales.Where(w => w.SaleNumber == dt.SaleNumber).ToList();
                List<DGSaleItems> tmpLst = new List<DGSaleItems>();
                foreach (var item in sales)
                {
                    if (dt.ProdId != item.ProductId && item.ProductId!=null)
                    {
                        //lst.Remove(item);
                        tmpLst.Add(item);//sales.Remove(item);
                    }
                    if (dt.ServId != item.ServiceId && item.ServiceId != null)
                    {
                        //lst.Remove(item);
                        tmpLst.Add(item);//sales.Remove(item);
                    }
                    
                }
                sales = tmpLst;
                salesToCheck = tmpLst;
                decimal totalSum = 0;
                if (sales.Any(a => a.ProdName == "Общая сумма чека"))
                {
                    var sale2 = sales.Single(s => s.ProdName == "Общая сумма чека");
                    sales.RemoveAt(sales.IndexOf(sale2));
                }
                foreach (var sale2 in sales)
                {
                    totalSum += sale2.Price;
                }
                if(totalSum>0)
                    sales.Add(new DGSaleItems() { ProdName = "Общая сумма чека", Price = totalSum });
                DGSaleFix.ItemsSource = sales;
                DGSaleFix.Items.Refresh();
            }
            catch (Exception ex)
            {
                int point = 0;
            }
        }
    }
}
