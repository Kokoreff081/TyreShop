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
    
    public partial class shop_review
    {
        public int review_id { get; set; }
        public int product_id { get; set; }
        public int customer_id { get; set; }
        public string author { get; set; }
        public string text { get; set; }
        public int rating { get; set; }
        public bool status { get; set; }
        public System.DateTime date_added { get; set; }
        public System.DateTime date_modified { get; set; }
    }
}