using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.IO;
using Bunifu.Framework.UI;

namespace CafeProject
{
    public partial class Register : UserControl
    {
        public Register()
        {
            InitializeComponent();
        }

        CMessageBox cMessageBox = new CMessageBox();

        static public BunifuFlatButton backbutton = new BunifuFlatButton();
        static public BunifuFlatButton backbutton2 = new BunifuFlatButton();
        static public BunifuDropdown regpositionscombobox = new BunifuDropdown();
        private void Register_Load(object sender, EventArgs e)
        {
            backbutton = BackBTN;
            backbutton2 = BackBTN2;
            regpositionscombobox = RegPositionsComboBox;

            //RegPositionsComboBox.AddItem("");
            //foreach (var item in Enum.GetValues(typeof(Positions)))
            //{
            //    RegPositionsComboBox.AddItem(item.ToString());
            //}

            //RegPositionsComboBox.RemoveAt(2);
        }

        private void RegBTN_Click(object sender, EventArgs e)
        {
            List<BunifuMaterialTextbox> boxes = new List<BunifuMaterialTextbox>()
            {
                RegNameBox, RegLastnameBox, RegUsernameBox, RegPasswordBox, RegVerificationBox
            };

            if (Rules.AlreadyRegistered(UsersDatabase.Base, RegUsernameBox, RegPasswordBox) == true)
            {
                cMessageBox.Show("ნიქნეიმი და პაროლი გამოყენებულია!");
            }

            if (Rules.Missedbox(boxes, RegPositionsComboBox) == false && Rules.AlreadyRegistered(UsersDatabase.Base, RegUsernameBox, RegPasswordBox) == false && Rules.MoreThanEight(RegPasswordBox, 8) == true && Rules.NeedSymbol(RegPasswordBox, 33, 64) == true && Rules.NeedSymbol(RegPasswordBox, 65, 90) == true && Rules.NeedSymbol(RegNameBox, 33, 64) == false && Rules.NeedSymbol(RegLastnameBox, 33, 64) == false && Rules.Verification(RegPasswordBox, RegVerificationBox) == true)
            {
                Rules.FirstUpper(RegNameBox);
                Rules.FirstUpper(RegLastnameBox);
                UsersDatabase.Base.Add(new UsersInfo { Name = RegNameBox.Text, Lastname = RegLastnameBox.Text, Username = RegUsernameBox.Text, Password = RegPasswordBox.Text, Position = RegPositionsComboBox.selectedValue.ToString() });
                cMessageBox.Show("რეგისტრაცია წარმატებით გაირეთ!");
                RegPositionsComboBox.selectedIndex = 0;
                foreach (var item in boxes)
                {
                    item.Text = null;
                }
               

                string Filename = "Database.txt";
                var json = JsonConvert.SerializeObject(UsersDatabase.Base);
                File.WriteAllText(Filename, json);


                using (StreamReader st = new StreamReader(Filename))
                {
                    string text = st.ReadToEnd();
                    var result = JsonConvert.DeserializeObject<BindingList<UsersInfo>>(text);
                }
            }
            else
            {
                cMessageBox.Show("რეგისტრაცია ვერ გაიარეთ!");
            }

            if (Rules.MoreThanEight(RegPasswordBox, 8) == false)
            {
                cMessageBox.Show("პაროლი უნდა იყოს რვა სიმბოლოზე მეტი!");
            }

            if (Rules.NeedSymbol(RegPasswordBox, 33, 65) == false)
            {
                cMessageBox.Show("პაროლში გჭირდება სიმბოლო!");
            }

            if (Rules.NeedSymbol(RegPasswordBox, 65, 90) == false)
            {
                cMessageBox.Show("პაროლში გჭირდება დიდი ასო!");
            }


            else if (Rules.Missedbox(boxes, RegPositionsComboBox) == true)
            {
                cMessageBox.Show("შეავსე ყველა ველი!");
            }

            if (Rules.NeedSymbol(RegNameBox, 33, 64) == true && RegNameBox.Text.Length > 0)
            {
                cMessageBox.Show("სახელს არ უნდა ქონდეს სიმბოლო!");
            }

            if (Rules.NeedSymbol(RegLastnameBox, 33, 64) == true && RegLastnameBox.Text.Length > 0)
            {
                cMessageBox.Show("გვარს არ უნდა ქონდეს სიმბოლო!");
            }

            if (Rules.Verification(RegPasswordBox, RegVerificationBox) == false)
            {
                cMessageBox.Show("პაროლმა შემოწმება ვერ გაიარა!");
            }
        }

        private void BackBTN_Click(object sender, EventArgs e)
        {
            this.Hide();
            MainWindow.form.Width = 282;
            MainWindow.form.Height = 358;
            FormsList.forms["RemoveOrCreate"].Show();
        }

        private void BackBTN2_Click(object sender, EventArgs e)
        {
            this.Hide();
            FormsList.forms["ManagerWindow"].Show();
            MainWindow.form.Width = 282;
            MainWindow.form.Height = 459;
        }

        private void RegPositionsComboBox_onItemSelected(object sender, EventArgs e)
        {

        }

        private void RegPasswordBox_OnValueChanged(object sender, EventArgs e)
        {
            RegPasswordBox.isPassword = true;
        }

        private void RegNameBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                RegBTN_Click(null, null);
            }
        }

        private void RegLastnameBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                RegBTN_Click(null, null);
            }
        }

        private void RegUsernameBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                RegBTN_Click(null, null);
            }
        }

        private void RegPasswordBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                RegBTN_Click(null, null);
            }
        }

        private void RegVerificationBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                RegBTN_Click(null, null);
            }
        }

        private void RegVerificationBox_OnValueChanged(object sender, EventArgs e)
        {
            RegVerificationBox.isPassword = true;
        }
    }
}
