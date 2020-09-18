using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CafeProject
{
    class Waitress
    {
        static public BindingList<BindingList<OrderItem>> TableList = new BindingList<BindingList<OrderItem>>();
        static public List<string> Descriptions = new List<string>();
        static public int Which = new int();
    }
}
