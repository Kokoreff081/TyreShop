//------------------------------------------------------------------------------
// <auto-generated>
//     Этот код создан по шаблону.
//
//     Изменения, вносимые в этот файл вручную, могут привести к непредвиденной работе приложения.
//     Изменения, вносимые в этот файл вручную, будут перезаписаны при повторном создании кода.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Tyreshop.DbAccess
{
    using System;
    using System.Collections.Generic;
    
    public partial class shop_product
    {
        public int product_id { get; set; }
        public string model { get; set; }
        public string sku { get; set; }
        public string upc { get; set; }
        public string ean { get; set; }
        public string jan { get; set; }
        public string isbn { get; set; }
        public string mpn { get; set; }
        public string location { get; set; }
        public int quantity { get; set; }
        public int stock_status_id { get; set; }
        public string image { get; set; }
        public int manufacturer_id { get; set; }
        public bool shipping { get; set; }
        public decimal price { get; set; }
        public int points { get; set; }
        public int tax_class_id { get; set; }
        public System.DateTime date_available { get; set; }
        public decimal weight { get; set; }
        public int weight_class_id { get; set; }
        public decimal length { get; set; }
        public decimal width { get; set; }
        public decimal height { get; set; }
        public int length_class_id { get; set; }
        public bool subtract { get; set; }
        public int minimum { get; set; }
        public int sort_order { get; set; }
        public bool status { get; set; }
        public int viewed { get; set; }
        public System.DateTime date_added { get; set; }
        public System.DateTime date_modified { get; set; }
    }
}