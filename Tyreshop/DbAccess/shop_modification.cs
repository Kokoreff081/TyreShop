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
    
    public partial class shop_modification
    {
        public int modification_id { get; set; }
        public string name { get; set; }
        public string code { get; set; }
        public string author { get; set; }
        public string version { get; set; }
        public string link { get; set; }
        public string xml { get; set; }
        public bool status { get; set; }
        public System.DateTime date_added { get; set; }
    }
}