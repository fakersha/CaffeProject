using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CafeProject
{
    public class ProductsItem
    {
        public string Name { get; set; }
        public int Count { get; set; }
        public double Weight { get; set; }
        public double Price { get; set; }

        public ProductsItem(string _name, double _weight, double _price)
        {
            this.Name = _name;
            this.Weight = _weight;
            this.Price = _price;
        }

        public ProductsItem(string _name, double _weight, double _price, int _count)
        {
            this.Name = _name;
            this.Count = _count;
            this.Weight = _weight;
            this.Price = _price;
        }

        ProductsItem()
        {

        }


    }
}
