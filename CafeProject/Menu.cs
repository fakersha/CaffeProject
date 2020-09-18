using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CafeProject
{
    class Menu
    {

        static public BindingList<MenuItem> MenuItems = new BindingList<MenuItem>();
        static public BindingList<MenuItem> CopyMenuItems = new BindingList<MenuItem>();

        static public void AddInMenu(MenuItem item)
        {
            MenuItems.Add(item);

        }
        static public void RemoveFromMenu(MenuItem item)
        {
            MenuItems.Remove(item);

        }


    }
}
