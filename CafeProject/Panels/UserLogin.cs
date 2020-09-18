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


namespace CafeProject
{
    public partial class UserLogin : UserControl
    {
        public UserLogin()
        {
            InitializeComponent();
            
        }

        CMessageBox cMessageBox = new CMessageBox();
        private void UserLogin_Load(object sender, EventArgs e)
        {
            string DatabaseFilename = "Database.txt";

            using (StreamReader st = new StreamReader(DatabaseFilename))
            {
                string text = st.ReadToEnd();
                var result = JsonConvert.DeserializeObject<BindingList<UsersInfo>>(text);
                UsersDatabase.Base = result;

            }


            string AccountantFilename = "AccountantProductBase.txt";

            using (StreamReader st = new StreamReader(AccountantFilename))
            {
                string text = st.ReadToEnd();
                var result = JsonConvert.DeserializeObject<BindingList<OrderItem>>(text);
                Accountant.AccountantList = result;

            }

            string readyfood = "ReadyFood.txt";

            using (StreamReader st = new StreamReader(readyfood))
            {
                string text = st.ReadToEnd();
                var result = JsonConvert.DeserializeObject<List<string>>(text);
                ResultBox.resultcombobox.Clear();
                foreach (var item in result)
                {
                    ResultBox.resultcombobox.AddItem(item);
                }
            }

            using (StreamReader st = new StreamReader("PlusNumber.txt"))
            {
                string text = st.ReadToEnd();
                if (text == "")
                {
                    text = "0";
                }
                var result = int.Parse(text);
                AccountantWindow.plus = result;

            }

            using (StreamReader st = new StreamReader("ImportedPlusNumber.txt"))
            {
                string text = st.ReadToEnd();
                if (text == "")
                {
                    text = "0";
                }
                var result = int.Parse(text);
                AddImportedProducts.plus = result;

            }

            using (StreamReader st = new StreamReader("MenuPlusNumber.txt"))
            {
                string text = st.ReadToEnd();
                if (text == "")
                {
                    text = "0";
                }
                var result = int.Parse(text);
                ManagerNewProducts.plus = result;

            }
        }

        private void AccountantSum()
        {
            List<OrderItem> b = Accountant.AccountantList
           .GroupBy(x => new { x.PayType, x.Name })
           .Select(x => new OrderItem { PayType = x.Key.PayType, Name = x.Key.Name, Count = x.Sum(y => y.Count), Price = x.Sum(y => y.Price), Type = x.First().Type })
           .ToList();

            Accountant.AccountantList.Clear();

            foreach (var item in b)
            {
                Accountant.AccountantList.Add(item);
            }
        }

        private void KitchenSum()
        {
            List<OrderForKitchen> b = KitchenTables.TableList[KitchenTables.Which]
           .GroupBy(x => x.Name)
           .Select(x => new OrderForKitchen(Name = x.Key, x.Sum(y => y.Count), x.First().Type))
           .ToList();

            KitchenTables.TableList[KitchenTables.Which].Clear();

            foreach (var item in b)
            {
                KitchenTables.TableList[KitchenTables.Which].Add(item);
            }
        }

        private void UserLogPasswordBox_OnValueChanged(object sender, EventArgs e)
        {
            UserLogPasswordBox.isPassword = true;
        }

        private void UserLoginBTN_Click(object sender, EventArgs e)
        {
            var Owners = Rules.PositionsLists(Positions.მეპატრონე);
            var Weitress = Rules.PositionsLists(Positions.მიმტანი);
            var Accountants = Rules.PositionsLists(Positions.ბუღალტერი);
            var Managers = Rules.PositionsLists(Positions.მენეჯერი);
            var Chefs = Rules.PositionsLists(Positions.მზარეული);
            bool login = false;
            bool incorrectpassword = false;
            foreach (var item in UsersDatabase.Base)
            {
                if (UserLogUsernameBox.Text == item.Username && UserLogPasswordBox.Text == item.Password)
                {
                    login = true;
                }

                if (UserLogUsernameBox.Text == item.Username && UserLogPasswordBox.Text != item.Password)
                {
                    incorrectpassword = true;
                }
            }

            if (Rules.PositionLogin(Owners, UserLogUsernameBox, UserLogPasswordBox) == true)
            {
                this.Hide();
                MainWindow.form.Width = 282;
                MainWindow.form.Height = 358;
                FormsList.forms["RemoveOrCreate"].Show();
                UserLogUsernameBox.Text = null;
                UserLogPasswordBox.Text = null;
                Register.backbutton2.Hide();
                Register.backbutton.Show();
                UsersRemover.backbutton2.Hide();
                UsersRemover.backbutton.Show();
                Register.regpositionscombobox.Clear();
                Register.regpositionscombobox.AddItem("");
                foreach (var item in Enum.GetValues(typeof(Positions)))
                {
                    Register.regpositionscombobox.AddItem(item.ToString());
                }

                Register.regpositionscombobox.RemoveAt(2);

            }
            else if (Rules.PositionLogin(Weitress, UserLogUsernameBox, UserLogPasswordBox) == true)
            {
                this.Hide();

                MainWindow.form.Width = 1207;
                MainWindow.form.Height = 533;
                FormsList.forms["WaiterWindow"].Show();
                UserLogUsernameBox.Text = null;
                UserLogPasswordBox.Text = null;
                if (ResultBox.resultcombobox.Items.Length == 0)
                {
                    ResultBox.resultcombobox.ForeColor = Color.Black;
                    ResultBox.resultcombobox.Text = "გამზადებული შეკვეთები!";
                }

            }
            else if (Rules.PositionLogin(Managers, UserLogUsernameBox, UserLogPasswordBox) == true)
            {
                this.Hide();

                MainWindow.form.Width = 282;
                MainWindow.form.Height = 459;
                FormsList.forms["ManagerWindow"].Show();
                UserLogUsernameBox.Text = null;
                UserLogPasswordBox.Text = null;
                UsersRemover.backbutton.Hide();
                UsersRemover.backbutton2.Show();
                Register.backbutton.Hide();
                Register.backbutton2.Show();
                Register.regpositionscombobox.Clear();
                Register.regpositionscombobox.AddItem("");
                foreach (var item in Enum.GetValues(typeof(Positions)))
                {
                    Register.regpositionscombobox.AddItem(item.ToString());
                }
                Register.regpositionscombobox.RemoveAt(1);
                Register.regpositionscombobox.RemoveAt(1);
                
            }
            else if (Rules.PositionLogin(Chefs, UserLogUsernameBox, UserLogPasswordBox) == true)
            {
                this.Hide();

                MainWindow.form.Width = 604;
                MainWindow.form.Height = 544;
                FormsList.forms["KitchenWindow"].Show();
                UserLogUsernameBox.Text = null;
                UserLogPasswordBox.Text = null;
                KitchenSum();
            }
            else if (Rules.PositionLogin(Accountants, UserLogUsernameBox, UserLogPasswordBox) == true)
            {
                this.Hide();

                MainWindow.form.Width = 621;
                MainWindow.form.Height = 572;
                FormsList.forms["AccountantWindow"].Show();
                string Filename = "AccountantProductBase.txt";
                var json = JsonConvert.SerializeObject(Accountant.AccountantList);
                File.WriteAllText(Filename, json);

                using (StreamReader st = new StreamReader(Filename))
                {
                    string text = st.ReadToEnd();
                    var result = JsonConvert.DeserializeObject<BindingList<OrderItem>>(text);
                    Accountant.AccountantList = result;
                    Gridviews.accountantgridview.DataSource = Accountant.AccountantList;
                }
                FormsList.forms["AccountantWindow"].Refresh();
                UserLogUsernameBox.Text = null;
                UserLogPasswordBox.Text = null;
                AccountantSum();
            }
            else if (login == true)
            {
                cMessageBox.Show("წარმატებით შეხვედით!");
                UserLogUsernameBox.Text = null;
                UserLogPasswordBox.Text = null;
            }
            else if (incorrectpassword == true && UserLogPasswordBox.Text.Length > 0)
            {
                cMessageBox.Show("პაროლი არასწორია!");
            }
            else if (UserLogUsernameBox.Text.Length == 0 || UserLogPasswordBox.Text.Length == 0)
            {
                cMessageBox.Show("შეავსე ყველა ველი!");
            }
            else
            {
                cMessageBox.Show("მომხმარებელი არ არის რეგისტრირებული!");
            }
        }

        private void UserLogUsernameBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                UserLoginBTN_Click(null, null);
            }
        }

        private void UserLogPasswordBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                UserLoginBTN_Click(null, null);
            }
        }
    }
}
