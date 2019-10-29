using Microsoft.Win32;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data.Entity.Validation;
using System.Drawing.Printing;
using System.Globalization;
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
    /// Логика взаимодействия для DailyReportWindow.xaml
    /// </summary>
    public partial class DailyReportWindow : Window
    {
        private List<DGSaleItems> lst;
        private List<PComboBox> list;
        public List<user> users;
        public DailyReportWindow(List<PComboBox> lstP)
        {
            InitializeComponent();
            lst = new List<DGSaleItems>();
            list = lstP;
            StartUpLoad();
        }

        private void StartUpLoad() {
            var today = DateTime.Now.Date;
            var time = DateTime.Now.TimeOfDay;
            string message = "За текущий день еще не было операций";
            GetGrid(today, list, message);
            
        }

        private void GetGrid(DateTime today, List<PComboBox>list, string msg) {
            using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
            {
                try
                {
                    var operations = db.operations.Where(w => w.OperationDate == today).ToList();
                    users = db.users.Where(w => w.Role == "manager").ToList();
                    //Report.Columns[]
                    decimal totalSum = 0;
                    foreach (var oper in operations)
                    {
                        string uName = "";
                        if (users.Exists(e => e.UserId == oper.UserId))
                            uName = users.Single(s => s.UserId == oper.UserId).UserName;
                        else
                            uName = users.Single(s => s.UserId == 14).UserName;
                        DelSaleButtonTag tag = new DelSaleButtonTag();
                        if (oper.ProductId!=null) {
                            tag.ServId = null;
                            tag.ProdId = oper.ProductId;
                            tag.AnotherId = null;
                            tag.SaleNumber = (long)oper.SaleNumber;
                        }
                        if (oper.ServiceId != null) {
                            tag.ServId = oper.ServiceId;
                            tag.ProdId = null;
                            tag.AnotherId = null;
                            tag.SaleNumber = (long)oper.SaleNumber;
                        }
                        DGSaleItems dg = new DGSaleItems();
                        dg.Comment = oper.Comment;
                        dg.Date = oper.OperationDate.ToString();
                        dg.OperationType = oper.OperationType;
                        dg.PayType = oper.PayType;
                        dg.Price = oper.Price;
                        if (oper.ProductId != null && oper.ProductId!=0)
                            dg.ProdName = list.Where(w => w.ProductId == oper.ProductId).Select(s => s.ProductName).First();
                        else if(oper.ServiceId!=null)
                            dg.ProdName = db.services.Where(w => w.ServiceId == oper.ServiceId).Select(s => s.ServiceName).First();
                        dg.Time = oper.OperationTime.ToString();
                        dg.Price = oper.Price;
                        dg.Quantity = oper.Count;
                        dg.StoreHouse = oper.Storehouse;
                        dg.SaleNumber = (long)oper.SaleNumber;
                        dg.CardToTotalSum = oper.CardToTotalSum;
                        dg.TagToBtn = tag;
                        dg.ProductId = oper.ProductId;
                        dg.ServiceId = oper.ServiceId;
                        if (oper.PayType == "Наличный расчет" && oper.OperationType != "Списание наличных")
                        {
                            if (oper.CardPay == "Да")
                            {
                                if (oper.CardToTotalSum != 0)
                                    totalSum += oper.Price;
                            }
                            else
                                totalSum += oper.Price;
                        }
                        if (oper.OperationType == "Списание наличных")
                            totalSum -= oper.Price;
                        dg.UserName = uName;
                        dg.User = users;
                        lst.Add(dg);
                    }
                    lst.Add(new DGSaleItems() { ProdName = "Сдача наличных", Price = totalSum });
                    Report.ItemsSource = lst;
                    Report.Items.Refresh();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(msg, "Информация", MessageBoxButton.OK);
                }
            }
        }

        private void SelectDate_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            DatePicker d = sender as DatePicker;
            var date = d.SelectedDate.Value.Date;
            var time = d.SelectedDate.Value.TimeOfDay;
            lst = new List<DGSaleItems>();
            string message = "Не найдено операций на выбранную дату!";
            GetGrid(date, list, message);
        }

        private void GenerateReport_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel file (*.xlsx)|*.xlsx";
            if (sfd.ShowDialog() == true)
            {
                try
                {
                    DateTime? date = null;
                    if ((DateTime?)SelectDate.SelectedDate == null)
                        date = DateTime.Now.Date;
                    else
                        date = (DateTime?)SelectDate.SelectedDate.Value.Date;
                    XlsxExport.ExportDailyReport(sfd, date);
                }
                catch (Exception ex)
                {
                    string msg = "Выбранный файл недоступен: " + ex.Message + "/r/nПожалуйста, повторите сохранение, задав другое имя файла.";
                    MessageBoxResult res = MessageBox.Show(msg, "Информация", MessageBoxButton.OK);
                }
            }
        }

        private void CashOff_Click(object sender, RoutedEventArgs e)
        {
            CultureInfo provider = CultureInfo.InvariantCulture;
            if (NumberCashOff.Text != string.Empty && CommentCashOff.Text != string.Empty) {
                decimal cash = 0;
                if (decimal.TryParse(NumberCashOff.Text, out var tmp)) {
                    cash = decimal.Parse(NumberCashOff.Text);
                }
                string comment = CommentCashOff.Text;
                
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities()) {
                    var sNum = db.operations.ToList();
                    long number = 0;
                    if (sNum.Count == 0)
                        number = 1;
                    else
                        number = (long)sNum[sNum.Count - 1].SaleNumber + 1;
                    var newSale = new operation();
                    newSale.SaleNumber = number;
                    newSale.PayType = "";
                    newSale.Comment = comment;
                    newSale.Price = cash;
                    newSale.OperationDate = DateTime.ParseExact(DateTime.Now.ToString("dd-MM-yyyy"), "dd-MM-yyyy", provider);
                    newSale.OperationTime = DateTime.ParseExact(DateTime.Now.ToString("hh:mm:ss"), "hh:mm:ss", provider);
                    newSale.OperationType = "Списание наличных";
                    newSale.Storehouse = "";
                    newSale.CardPay = "Нет";
                    newSale.ProductId = null;
                    newSale.ServiceId = null;
                    newSale.Count = 0;
                    try
                    {
                        db.operations.Add(newSale);
                        db.SaveChanges();
                    }
                    catch(DbEntityValidationException ex)
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
                    var today = DateTime.Now.Date;
                    var list = db.products.Select(s => new PComboBox()
                    {
                        ProductId = s.ProductId,
                        ProductName = @"" + db.manufacturers.Where(w => w.ManufacturerId == s.ManufacturerId).Select(sel => sel.ManufacturerName).FirstOrDefault() + " " +
                    db.models.Where(w => w.ModelId == s.ModelId && w.ManufacturerId == s.ManufacturerId).Select(sel => sel.ModelName).FirstOrDefault() + " " + s.Width + " / " + s.Height + " / " + s.Radius
                    }).ToList();
                    MessageBox.Show("Наличные успешно списаны!", "Информация", MessageBoxButton.OK);
                    GetGrid(today, list, "");
                }
            }
        }

        private void DelSale_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            DelSaleButtonTag dt = btn.Tag as DelSaleButtonTag;
            try
            {
                DGSaleItems sale=new DGSaleItems();
                var sales = lst.Where(w => w.SaleNumber == dt.SaleNumber).ToList();
                foreach (var item in sales)
                {
                    if (dt.ProdId != null)
                    {
                        sale = lst.First(s => s.ProductId == dt.ProdId && s.SaleNumber == dt.SaleNumber);
                    }
                    if (dt.ServId != null)
                    {
                        sale = lst.First(s => s.ServiceId == dt.ServId && s.SaleNumber == dt.SaleNumber);
                    }
                    lst.Remove(sale);
                    using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                    {
                        var oper = db.operations.First(s => (s.ProductId == dt.ProdId && s.SaleNumber == dt.SaleNumber) || (s.ServiceId == dt.ServId && s.SaleNumber == dt.SaleNumber));
                        var res = MessageBox.Show("Вы действительно хотите полностью удалить данную операцию? Действие необратимо.", "Информация", MessageBoxButton.OKCancel);
                        if (res == MessageBoxResult.OK)
                        {
                            if (oper.ProductId != null && oper.ProductId != 0)
                            {
                                int store = db.storehouses.First(s => s.StorehouseName == oper.Storehouse).StorehouseId;
                                int quant = oper.Count;
                                var pq = db.productquantities.First(s => s.StorehouseId == store && s.ProductId == oper.ProductId);
                                pq.Quantity += quant;
                                db.Entry(pq).Property(x => x.Quantity).IsModified = true;
                            }

                            db.operations.Remove(oper);
                            db.SaveChanges();
                        }
                    }
                }
                Report.Items.Refresh();
            }
            catch (Exception ex) {
                int point = 0;
            }
        }

        private void ManagerCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var cmb = sender as ComboBox;
            var sNum = (long)cmb.Tag;
            try
            {
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities()) {
                    var operations = db.operations.Where(w => w.SaleNumber == sNum).ToList();
                    foreach (var oper in operations) {
                        oper.UserId = (int)cmb.SelectedValue;
                        db.Entry(oper).Property(x => x.UserId).IsModified = true;
                        db.SaveChanges();
                    }
                }
            }
            catch (Exception ex) {
                int point = 0;
            }
        }

        private void PrintSale_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            DelSaleButtonTag dt = btn.Tag as DelSaleButtonTag;
            try {

                string saleNumber = "";
                string saleDate = "";
                string productPrice = "";
                List<operation> saleNumbers = new List<operation>();
                var tmp = lst.Where(w => w.SaleNumber == dt.SaleNumber).ToList();
                List<DGSaleItems> sales = new List<DGSaleItems>();
                foreach (var t in tmp) {
                    if (!sales.Any(a => a.ProductId == t.ProductId))
                        sales.Add(t);
                    else
                    {
                        var sale = sales.Single(w => w.ProductId == t.ProductId);
                        sale.Quantity += t.Quantity;
                        sale.Price += t.Price;
                    }
                }
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    saleNumbers = db.operations.Where(w=>w.SaleNumber==dt.SaleNumber).ToList();
                    saleNumber = dt.SaleNumber.ToString();
                    saleDate = saleNumbers[0].OperationDate.ToString("m", CultureInfo.CurrentUICulture) + " " + saleNumbers[0].OperationDate.ToString("yyyy");
                    var pId = saleNumbers[0].ProductId;
                    productPrice = db.products.Where(w => w.ProductId == pId).Select(s => s.Price).FirstOrDefault().ToString();
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
                sheet.Range["C9"].Text = "Товарный чек № " + saleNumber + " от ";
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
                for (int i = 0; i < sales.Count; i++)
                {
                    var item = sales[i];
                    sheet.Range["A" + counter + ":B" + counter].Merge();
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
                    sheet.Range["M" + counter].Text = item.Price.ToString() + " руб.";
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
                sheet.Range["C" + counter + ":K" + (counter + 4)].Merge();
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
            catch (Exception ex)
            {
                int point = 0;
            }
        }

        private void GenerateReportPeriod_Click(object sender, RoutedEventArgs e)
        {
            if (SelectDateFrom.SelectedDate != null && SelectDateTo.SelectedDate != null) {
                var from = SelectDateFrom.SelectedDate;
                var to = SelectDateTo.SelectedDate;
                try
                {
                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Filter = "Excel file (*.xlsx)|*.xlsx";
                    if (sfd.ShowDialog() == true)
                    {
                        XlsxExport.ExportDailyReport(sfd, from, to);
                    }
                }
                catch (Exception ex) {
                    int point = 0;
                }
            }
        }
    }
}
