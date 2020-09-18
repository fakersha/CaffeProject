using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CafeProject
{
    class Accountant
    {
        static public BindingList<OrderItem> AccountantList = new BindingList<OrderItem>();
        static public BindingList<OrderItem> CopyAccountantList = new BindingList<OrderItem>();
        static public BindingList<OrderItem> PayCopyAccountantList = new BindingList<OrderItem>();
    }
}
