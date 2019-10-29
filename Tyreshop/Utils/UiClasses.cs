using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tyreshop.DbAccess;
using NPOI;
using Ganss.Excel;

namespace Tyreshop.Utils
{
    class TreeViewContent:INotifyPropertyChanged
    {
        private int _categoryId;
        public int CategoryId { get { return _categoryId; } set { _categoryId = value; OnPropertyChanged("CategoryId"); } }
        private string _categoryName;
        public string CategoryName { get { return _categoryName; } set { _categoryName = value; OnPropertyChanged("CategoryName"); } }
        public List<GroupOfProducts> Products { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged(string prop)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }
    }

    class GroupOfProducts:INotifyPropertyChanged
    {
        private string _radius;
        public string Radius { get { return _radius; } set { _radius = value; OnPropertyChanged("Radius"); } }
        private int storeId;
        public int StoreId { get { return storeId; } set { storeId = value; OnPropertyChanged("StoreId"); } }
        private string storeName;
        public string StoreName { get { return storeName; } set { storeName = value; OnPropertyChanged("StoreName"); } }
        public List<BdProducts> BdProducts { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged(string prop)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }
    }

    public class BdProducts : INotifyPropertyChanged
    {
        public int ProductId { get; set; }
        public int ModelId { get; set; }
        public int CategoryId { get; set; }
        public int Width { get; set; }
        public float Height { get; set; }
        public string Radius { get; set; }
        public string Manufacturer { get; set; }
        public string Model { get; set; }
        public string Articul { get; set; }
        public string Season { get; set; }
        public string InCol { get; set; }
        public string IsCol { get; set; }
        public string Gruz { get; set; }
        public string RFT { get; set; }
        public string Spikes { get; set; }
        public decimal Price { get; set; }
        public decimal? OptPrice { get; set; }
        public decimal? PurPrice { get; set; }
        public List<ProductQuantity> Storehouse { get; set; }
        public int? TotalQuantity { get; set; }
        public string StorehouseName { get; set; }
        public string Address { get; set; }
        public int StorehouseId { get; set; }
        public int? InOrder { get; set; }
        public string Country { get; set; }
        public decimal? ComplektPrice { get; set; }


        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged(string prop)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }
    }

    public class ProductQuantity 
    {
        public int Id { get; set; }
        public string StorehouseName { get; set; }
        public int StorehouseId { get; set; }
        public string Address { get; set; }
        public int ProductId { get; set; }
        public int? Quantity { get; set; }
        public int? InOrder { get; set; }
        
    }

    public class Manufacturers {
        public int ProductId { get; set; }
        public string Manufacturer { get; set; }
    }

    public class Season {
        public int Id { get; set; }
        public string Name { get; set; }
    }

    public class RFT
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }

    public class Gruz
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }

    public class Spike
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }

    public class PComboBox
    {
        public int ProductId { get; set; }
        public string ProductName { get; set; }
    }

    public class DGSaleItems
    {
        public long SaleNumber { get; set; }
        public string ProdName { get; set; }
        public string Date { get; set; }
        public string Time { get; set; }
        public int Quantity { get; set; }
        public decimal Price { get; set; }
        public int? ProductId { get; set; }
        public int? ServiceId { get; set; }
        public string OperationType { get; set; }
        public string PayType { get; set; }
        public int? StoreId { get; set; }
        public string Comment { get; set; }
        public string StoreHouse { get; set; }
        public string CardPayed { get; set; }
        public int CardToTotalSum { get; set; }
        public DelSaleButtonTag TagToBtn { get; set; }
        public int? AnotherId { get; set; }
        public int AnotherQuant { get; set; }
        public string AnotherName { get; set; }
        public List<user> User { get; set; }
        public int UserId { get; set; }
        public string UserName { get; set; }
    }

    public class GridCapitalize {
        public int ProductId { get; set; }
        public string ProdName { get; set; }
        public string StorehouseName { get; set; }
        public int TotalQuantity { get; set; }
    }

    public class StoreQuant {
        public string PSeason { get; set; }
        public int PQuant { get; set; }
        public int TotalQuant { get; set; }
    }

    public class OrderList {
        public long OrderId { get; set; }
        public long OrderNumber { get; set; } 
        public string OrderCustomer { get; set; }
        public string PhoneCustomer { get; set; }
        public string ProdName { get; set; }
        public int Quantity { get; set; }
        public string Storehouse { get; set; }
        public string UserRole { get; set; }
        public string UserName { get; set; }
    }

    public class XlsxImport {
        public int ProductId { get; set; }
        public int CategoryId { get; set; }
        public int Width { get; set; }
        public float Height { get; set; }
        public string Radius { get; set; }
        public int Manufacturer { get; set; }
        public int Model { get; set; }
        public string Articul { get; set; }
        public string Season { get; set; }
        public string InCol { get; set; }
        public string IsCol { get; set; }
        public string Gruz { get; set; }
        public string RFT { get; set; }
        public string Spikes { get; set; }
        public decimal Price { get; set; }
        public decimal? OptPrice { get; set; }
        public decimal? PurPrice { get; set; }
        public List<ProductQuantity> Storehouse { get; set; }
        public int? TotalQuantity { get; set; }
        public string StorehouseName { get; set; }
        public string Address { get; set; }
        public int StorehouseId { get; set; }
        public int? InOrder { get; set; }
    }

    public class DelSaleButtonTag {
        public int? ProdId { get; set; }
        public int? ServId { get; set; }
        public int? AnotherId { get; set; }
        public long SaleNumber { get; set; }
    }

    public class bgWork {
        public string fileName { get; set; }
    }
}
