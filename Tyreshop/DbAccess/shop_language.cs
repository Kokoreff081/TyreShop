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
    
    public partial class shop_language
    {
        public int language_id { get; set; }
        public string name { get; set; }
        public string code { get; set; }
        public string locale { get; set; }
        public string image { get; set; }
        public string directory { get; set; }
        public int sort_order { get; set; }
        public bool status { get; set; }
    }
}