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
    
    public partial class shop_customer
    {
        public int customer_id { get; set; }
        public int customer_group_id { get; set; }
        public int store_id { get; set; }
        public string firstname { get; set; }
        public string lastname { get; set; }
        public string email { get; set; }
        public string telephone { get; set; }
        public string fax { get; set; }
        public string password { get; set; }
        public string salt { get; set; }
        public string cart { get; set; }
        public string wishlist { get; set; }
        public bool newsletter { get; set; }
        public int address_id { get; set; }
        public string custom_field { get; set; }
        public string ip { get; set; }
        public bool status { get; set; }
        public bool approved { get; set; }
        public bool safe { get; set; }
        public string token { get; set; }
        public System.DateTime date_added { get; set; }
    }
}
