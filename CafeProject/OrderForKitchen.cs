using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CafeProject
{
    public class OrderForKitchen
    {
        public string Name { get; set; }
        public int Count { get; set; }
        public string Type { get; set; }
        public OrderForKitchen(string _Name, int _orderCount, string _type)
        {
            this.Name = _Name;
            this.Count = _orderCount;
            this.Type = _type;
        }
    }
}
