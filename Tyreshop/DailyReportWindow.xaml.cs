using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data.Entity.Validation;
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
                        dg.UserName = users.Where(w => w.UserId == oper.UserId || w.UserId==14).Select(s => s.UserName).First();
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
                if (dt.ProdId != null)
                {
                    sale = lst.Single(s => s.ProductId == dt.ProdId && s.SaleNumber == dt.SaleNumber);
                }
                if (dt.ServId != null) {
                    sale = lst.Single(s=>s.ServiceId == dt.ServId && s.SaleNumber == dt.SaleNumber);
                }
                lst.Remove(sale);
                using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities())
                {
                    var oper = db.operations.Single(s => (s.ProductId == dt.ProdId && s.SaleNumber == dt.SaleNumber) || (s.ServiceId == dt.ServId && s.SaleNumber == dt.SaleNumber) );
                    var res = MessageBox.Show("Вы действительно хотите полностью удалить данную операцию? Действие необратимо.", "Информация", MessageBoxButton.OKCancel);
                    if (res == MessageBoxResult.OK)
                    {
                        if (oper.ProductId != null && oper.ProductId!=0)
                        {
                            int store = db.storehouses.Single(s => s.StorehouseName == oper.Storehouse).StorehouseId;
                            int quant = oper.Count;
                            var pq = db.productquantities.Single(s => s.StorehouseId == store && s.ProductId == oper.ProductId);
                            pq.Quantity += quant;
                            db.Entry(pq).Property(x => x.Quantity).IsModified = true;
                        }

                        db.operations.Remove(oper);
                        db.SaveChanges();
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
    }
}
