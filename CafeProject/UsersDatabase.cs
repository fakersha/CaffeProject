using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CafeProject
{
    class UsersDatabase
    {
        static public BindingList<UsersInfo> Base = new BindingList<UsersInfo>()
        {
            new UsersInfo {Name = "Giorgi", Lastname = "Gagnidze", Username = "Owner", Password = "Owner!", Position = Positions.მეპატრონე.ToString()}
        };
    }
}
