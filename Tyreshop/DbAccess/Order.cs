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
    
    public partial class Order
    {
        public long OrderID { get; set; }
        public string OrderCustomer { get; set; }
        public string PhoneCustomer { get; set; }
        public long OrderNumber { get; set; }
        public int StorehouseId { get; set; }
        public int ProductId { get; set; }
        public int ProductQuant { get; set; }
        public int MengerId { get; set; }
    }
}