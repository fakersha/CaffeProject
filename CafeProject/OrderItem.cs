using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CafeProject
{
    class OrderItem
    {
        public string Name { get; set; }
        public double Price { get; set; }
        public int Count { get; set; }
        public string Type { get; set; }
        public string PayType { get; set; }


        public OrderItem(string _Name, double _Price, int _orderCount, string _type, string _paytype)
        {
            this.Name = _Name;
            this.Price = _Price;
            this.Count = _orderCount;
            this.Type = _type;
            this.PayType = _paytype;
        }

        public OrderItem()
        {

        }
    }
}
