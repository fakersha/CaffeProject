using Bunifu.Framework.UI;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CafeProject
{
    class Rules
    {

        public static void FirstUpper(BunifuMaterialTextbox textBox)
        {
            string Upper = textBox.Text.First().ToString().ToUpper();
            string FirstRemove = textBox.Text.Remove(0, 1);
            textBox.Text = Upper + FirstRemove;
        }
        public static IEnumerable<UsersInfo> PositionsLists(Positions pos)
        {

            var roles = from role in UsersDatabase.Base
                        where role.Position == pos.ToString()
                        select role;

            return roles;

        }
        public static bool PositionLogin(IEnumerable<UsersInfo> list, BunifuMaterialTextbox textBox1, BunifuMaterialTextbox textBox2)
        {
            bool login = false;

            foreach (var item in list)
            {
                if (textBox1.Text == item.Username && textBox2.Text == item.Password)
                {
                    login = true;
                }
            }

            return login;
        }
        public static bool MoreThanEight(BunifuMaterialTextbox textBox, int numb)
        {
            bool eightsymbol = true;
            if (textBox.Text.Length < numb && textBox.Text.Length > 0)
            {
                eightsymbol = false;
            }
            return eightsymbol;
        }


        static public bool NeedSymbol(BunifuMaterialTextbox textBox, int first, int last)
        {
            bool symbol = false;
            for (int i = first; i < last + 1; i++)
            {
                if (textBox.Text.Contains((char)i) == true || textBox.Text.Length == 0)
                {
                    symbol = true;
                }
            }

            return symbol;
        }


        public static bool AlreadyRegistered(BindingList<UsersInfo> accounts, BunifuMaterialTextbox textBoxUsername, BunifuMaterialTextbox textBoxPassword)
        {
            bool already = false;

            foreach (var item in accounts)
            {
                if (item.Username == textBoxUsername.Text && item.Password == textBoxPassword.Text)
                {
                    already = true;
                }
            }
            return already;

        }


        public static bool Missedbox(List<BunifuMaterialTextbox> textBoxes, BunifuDropdown comboBox)
        {
            bool missed = false;
            
            foreach (var item in textBoxes)
            {
                try
                {
                    if (item.Text.Length == 0 && comboBox.selectedValue == "")
                    {

                    }
                    else if (item.Text.Length == 0)
                    {
                        missed = true;
                    }
                }
                catch (Exception)
                {
                    missed = true;
                }
                

            }

            return missed;
        }

        public static bool Verification(BunifuMaterialTextbox box1, BunifuMaterialTextbox box2)
        {
            bool verification = true;
            if (box1.Text != box2.Text)
            {
                verification = false;
            }

            return verification;
        }
    }
}
