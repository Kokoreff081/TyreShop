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
    
    public partial class shop_customer_reward
    {
        public int customer_reward_id { get; set; }
        public int customer_id { get; set; }
        public int order_id { get; set; }
        public string description { get; set; }
        public int points { get; set; }
        public System.DateTime date_added { get; set; }
    }
}
