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
    
    public partial class shop_recurring
    {
        public int recurring_id { get; set; }
        public decimal price { get; set; }
        public string frequency { get; set; }
        public long duration { get; set; }
        public long cycle { get; set; }
        public sbyte trial_status { get; set; }
        public decimal trial_price { get; set; }
        public string trial_frequency { get; set; }
        public long trial_duration { get; set; }
        public long trial_cycle { get; set; }
        public sbyte status { get; set; }
        public int sort_order { get; set; }
    }
}