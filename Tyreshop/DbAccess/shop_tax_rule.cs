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
    
    public partial class shop_tax_rule
    {
        public int tax_rule_id { get; set; }
        public int tax_class_id { get; set; }
        public int tax_rate_id { get; set; }
        public string based { get; set; }
        public int priority { get; set; }
    }
}
