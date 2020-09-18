using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CafeProject
{
    class MenuItem
    {
        public string Name { get; set; }
        public double Price { get; set; }
        public string Type { get; set; }

        public int ID { get; set; }


        public MenuItem(string _Name, double _Price, string _Type, int _ID)
        {
            this.Name = _Name;
            this.Price = _Price;
            this.Type = _Type;
            this.ID = _ID;
        }
        public MenuItem()
        {


        }
    }
}
