using Bunifu.Framework.UI;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CafeProject
{
    class KitchenTables
    {
        static public BindingList<BindingList<OrderForKitchen>> TableList = new BindingList<BindingList<OrderForKitchen>>()
       {
                new BindingList<OrderForKitchen>(),
                new BindingList<OrderForKitchen>(),
                new BindingList<OrderForKitchen>(),
                new BindingList<OrderForKitchen>(),
                new BindingList<OrderForKitchen>(),
                new BindingList<OrderForKitchen>(),
                new BindingList<OrderForKitchen>(),
                new BindingList<OrderForKitchen>()
       };

        static public int Which = new int();

        static public List<BunifuFlatButton> tables = new List<BunifuFlatButton>();

        static public List<string> readyfood = new List<string>();


    }
}
