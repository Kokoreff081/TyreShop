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
    
    public partial class shop_api_session
    {
        public int api_session_id { get; set; }
        public int api_id { get; set; }
        public string token { get; set; }
        public string session_id { get; set; }
        public string session_name { get; set; }
        public string ip { get; set; }
        public System.DateTime date_added { get; set; }
        public System.DateTime date_modified { get; set; }
    }
}