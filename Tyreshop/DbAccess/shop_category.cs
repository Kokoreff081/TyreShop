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
    
    public partial class shop_category
    {
        public int category_id { get; set; }
        public string image { get; set; }
        public int parent_id { get; set; }
        public bool top { get; set; }
        public int column { get; set; }
        public int sort_order { get; set; }
        public bool status { get; set; }
        public System.DateTime date_added { get; set; }
        public System.DateTime date_modified { get; set; }
    }
}
