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
    
    public partial class shop_product_special
    {
        public int product_special_id { get; set; }
        public int product_id { get; set; }
        public int customer_group_id { get; set; }
        public int priority { get; set; }
        public decimal price { get; set; }
        public System.DateTime date_start { get; set; }
        public System.DateTime date_end { get; set; }
    }
}
