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
    
    public partial class shop_cart
    {
        public int cart_id { get; set; }
        public int customer_id { get; set; }
        public string session_id { get; set; }
        public int product_id { get; set; }
        public string gift_teaser { get; set; }
        public int recurring_id { get; set; }
        public string option { get; set; }
        public int quantity { get; set; }
        public System.DateTime date_added { get; set; }
    }
}