using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CafeProject
{
    public class MenuT
    {
        public string Name { get; set; }
        public double Price { get; set; }
        public int OrderCount { get; set; }

        public MenuT(string _Name, double _Price)
        {
            this.Name = _Name;
            this.Price = _Price;
            //test

        }
        public MenuT(string _Name, double _Price, int _Ordercount)
        {
            this.Name = _Name;
            this.Price = _Price;
            this.OrderCount = _Ordercount;
            //test

        }

        public MenuT()
        {


        }
    }
}
