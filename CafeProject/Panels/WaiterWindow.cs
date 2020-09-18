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
    public partial class WaiterWindow : UserControl
    {
        private BindingList<OrderItem> OrderList;
        private List<BunifuFlatButton> tables = new List<BunifuFlatButton>();
        double sum = 0;
        public WaiterWindow()
        {
            InitializeComponent();
            OrderList = new BindingList<OrderItem>();
            Waitress.TableList = new BindingList<BindingList<OrderItem>>()
            {
                new BindingList<OrderItem>(),
                new BindingList<OrderItem>(),
                new BindingList<OrderItem>(),
                new BindingList<OrderItem>(),
                new BindingList<OrderItem>(),
                new BindingList<OrderItem>(),
                new BindingList<OrderItem>(),
                new BindingList<OrderItem>()
            };

            Waitress.Descriptions = new List<string>()
            {
                "", "", "", "", "", "", "", ""
            };

            tables = new List<BunifuFlatButton>()
            {
                 Table_1, Table_2, Table_3, Table_4, Table_5,Table_6,Table_7,Table_8
            };
            Menu.AddInMenu(new MenuItem("Margarita", 5, ProductType.პიცა.ToString(), 1111));
            Menu.AddInMenu(new MenuItem("Cola", 1.5, ProductType.სასმელები.ToString(), 2222));
            Menu.AddInMenu(new MenuItem("XL", 2.5, ProductType.სასმელები.ToString(), 3333));
            Menu.AddInMenu(new MenuItem("Cheese_burger", 5, ProductType.ბურგერი.ToString(), 4444));
        }

        CMessageBox cMessageBox = new CMessageBox();

        static public DataGridViewRow row = new DataGridViewRow();
        private void WaiterWindow_Load(object sender, EventArgs e)
        {
            ResultBox.resultcombobox = ResultCombobox;
            dataGridView1.DataSource = Menu.MenuItems;
            dataGridView1.Refresh();
            dataGridView3.Refresh();
            foreach (var item in Enum.GetValues(typeof(ProductType)))
            {
                TypeBox.AddItem(item.ToString());
            }
            foreach (var item in Enum.GetValues(typeof(PayTypes)))
            {
                PayTypeComboBox.AddItem(item.ToString());
            }
            PayTypeComboBox.selectedIndex = 0;
            TypeBox.selectedIndex = 0;
            
        }

        private void Delete_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell item in dataGridView3.SelectedCells)
            {
                try
                {
                    if (Waitress.TableList[Waitress.Which][item.RowIndex].Count > 0)
                    {
                        Waitress.TableList[Waitress.Which][item.RowIndex].Count--;
                        dataGridView3.Refresh();
                    }
                }
                catch (Exception)
                {

                }

                try
                {
                    if (Waitress.TableList[Waitress.Which][item.RowIndex].Count < 1)
                    {

                        dataGridView3.Rows.RemoveAt(item.RowIndex);
                        dataGridView3.Refresh();

                    }
                }
                catch (Exception)
                {

                }

            }

            WhenTabelClearColor();
        }


        private void WhenTabelClearColor()
        {
            if (Waitress.TableList[Waitress.Which].Count == 0)
            {
                tables[Waitress.Which].Normalcolor = Color.FromArgb(220, 151, 60, 49);
                tables[Waitress.Which].OnHovercolor = Color.FromArgb(138, 2, 2);
                tables[Waitress.Which].Activecolor = Color.FromArgb(138, 2, 2);
            }
        }
        private void RemoveClick_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell item in dataGridView3.SelectedCells)
            {
                try
                {
                    dataGridView3.Rows.RemoveAt(item.RowIndex);
                }
                catch (Exception)
                {

                }
                dataGridView3.Refresh();
            }
            WhenTabelClearColor();
        }

        private void sumOfOrder_Click(object sender, EventArgs e)
        {
            foreach (var item in Waitress.TableList[Waitress.Which])
            {
                sum += item.Price;
            }
            cMessageBox.Show($" შეკვეთის ღირებულება :{sum.ToString()} ₾");
            sum = 0;
        }

        private void TypeBox_onItemSelected(object sender, EventArgs e)
        {
            Menu.CopyMenuItems = new BindingList<MenuItem>();
            foreach (var item in Menu.MenuItems)
            {
                if (TypeBox.selectedValue.ToString() == item.Type.ToString())
                {
                    Menu.CopyMenuItems.Add(item);
                }
            }

            if (TypeBox.selectedValue.ToString() == ProductType.მენიუ.ToString())
            {
                dataGridView1.DataSource = Menu.MenuItems;
            }
            else
            {
                dataGridView1.DataSource = Menu.CopyMenuItems;
            }
            dataGridView1.Refresh();
        }

        private void ResultCombobox_onItemSelected(object sender, EventArgs e)
        {
            ResultCombobox.RemoveAt(ResultCombobox.selectedIndex);
        }

        private void PayTypeComboBox_onItemSelected(object sender, EventArgs e)
        {

        }

        private void TableClearBTN_Click(object sender, EventArgs e)
        {
            Waitress.TableList[Waitress.Which].Clear();

            WhenTabelClearColor();

            dataGridView3.Refresh();
        }

        private void Table_1_Click(object sender, EventArgs e)
        {
            dataGridView3.DataSource = Waitress.TableList[0];
            Waitress.Which = 0;
            BTNColors();
            dataGridView3.Refresh();
        }

        private void Table_2_Click(object sender, EventArgs e)
        {
            dataGridView3.DataSource = Waitress.TableList[1];
            Waitress.Which = 1;
            BTNColors();
            dataGridView3.Refresh();
        }

        private void Table_3_Click(object sender, EventArgs e)
        {
            dataGridView3.DataSource = Waitress.TableList[2];
            Waitress.Which = 2;
            BTNColors();
            dataGridView3.Refresh();
        }

        private void Table_4_Click(object sender, EventArgs e)
        {
            dataGridView3.DataSource = Waitress.TableList[3];
            Waitress.Which = 3;
            BTNColors();
            dataGridView3.Refresh();
        }

        private void Table_5_Click(object sender, EventArgs e)
        {
            dataGridView3.DataSource = Waitress.TableList[4];
            Waitress.Which = 4;
            BTNColors();
            dataGridView3.Refresh();
        }

        private void Table_6_Click(object sender, EventArgs e)
        {
            dataGridView3.DataSource = Waitress.TableList[5];
            Waitress.Which = 5;
            BTNColors();
            dataGridView3.Refresh();
        }

        private void Table_7_Click(object sender, EventArgs e)
        {
            dataGridView3.DataSource = Waitress.TableList[6];
            Waitress.Which = 6;
            BTNColors();
            dataGridView3.Refresh();
        }

        private void Table_8_Click(object sender, EventArgs e)
        {
            dataGridView3.DataSource = Waitress.TableList[7];
            Waitress.Which = 7;
            BTNColors();
            dataGridView3.Refresh();
        }

        private void Save_Click(object sender, EventArgs e)
        {
            foreach (var item in Waitress.TableList[Waitress.Which])
            {
                Accountant.AccountantList.Add(item);
                KitchenTables.TableList[Waitress.Which].Add(new OrderForKitchen(item.Name, item.Count, item.Type));
            }

            string Filename = "AccountantProductBase.txt";
            var json = JsonConvert.SerializeObject(Accountant.AccountantList);
            File.WriteAllText(Filename, json);

            using (StreamReader st = new StreamReader(Filename))
            {
                string text = st.ReadToEnd();
                var result = JsonConvert.DeserializeObject<BindingList<OrderItem>>(text);
                Accountant.AccountantList = result;

            }

            foreach (var item in KitchenTables.TableList)
            {
                if (item.Count > 0)
                {
                    KitchenTables.tables[KitchenTables.TableList.IndexOf(item)].Normalcolor = Color.FromArgb(247, 227, 96, 40);
                    KitchenTables.tables[KitchenTables.TableList.IndexOf(item)].OnHovercolor = Color.Orange;
                }
            }

            Waitress.Descriptions[Waitress.Which] = DescriptionBox.Text;
            DescriptionBox.Text = null;
        }

        private void BackBTN_Click(object sender, EventArgs e)
        {
            this.Hide();
            MainWindow.form.Width = 353;
            MainWindow.form.Height = 473;
            FormsList.forms["UserLogin"].Show();
        }

        private void BTNColors()
        {
            if (tables[Waitress.Which].Normalcolor != Color.FromArgb(247, 227, 96, 40))
            {
                tables[Waitress.Which].Normalcolor = Color.FromArgb(138, 2, 2);
            }
            foreach (var item in tables)
            {
                if (item != tables[Waitress.Which] && item.Normalcolor != Color.FromArgb(247, 227, 96, 40))
                {
                    item.Normalcolor = Color.FromArgb(220, 151, 60, 49);
                }
            }
        }
        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            row = this.dataGridView1.Rows[e.RowIndex];

            string RName = row.Cells["Name"].Value?.ToString();
            string RPrice = row.Cells["Price"].Value.ToString();
            double OrderPrice = double.Parse(RPrice);
            string RType = row.Cells["Type"].Value?.ToString();

            MenuItem orderItem = new MenuItem { Name = RName, Price = OrderPrice, Type = RType };


            if (Waitress.TableList[Waitress.Which].Count >= 0)
            {
                tables[Waitress.Which].Normalcolor = Color.FromArgb(247, 227, 96, 40);
                tables[Waitress.Which].OnHovercolor = Color.Orange;
                tables[Waitress.Which].Activecolor = Color.Orange;
            }



            try
            {
                if (RName != null && RType != null)
                {
                    Waitress.TableList[Waitress.Which].Add((new OrderItem { Name = RName, Price = OrderPrice, Count = 1, Type = RType, PayType = PayTypeComboBox.selectedValue.ToString() }));
                }
            }
            catch (Exception)
            {
                cMessageBox.Show("მიუთითეთ გადახდის ტიპი!");
            }
            
            List<OrderItem> b = Waitress.TableList[Waitress.Which].GroupBy(x => new { x.PayType, x.Name }).Select(x => new OrderItem { PayType = x.Key.PayType, Name = x.Key.Name, Count = x.Sum(y => y.Count), Price = x.Sum(y => y.Price), Type = x.First().Type })
           .ToList();

            Waitress.TableList[Waitress.Which].Clear();
            foreach (var item in b)
            {
                Waitress.TableList[Waitress.Which].Add(item);
            }





            dataGridView3.Refresh();
        }

        private void ChangeBox_KeyDown(object sender, KeyEventArgs e)
        {
            double change = new double();

            foreach (var item in Waitress.TableList[Waitress.Which])
            {
                change += item.Price;
            }

            double money = new double();


            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    money = double.Parse(ChangeBox.Text);
                    double result = money - change;
                    cMessageBox.Show(($"ხურდა: {result}₾"));
                    ChangeBox.Text = null;
                }
                catch (Exception)
                {
                    cMessageBox.Show("მნიშვნელობა უნდა იყოს რიცხვი!");
                }


            }
        }
    }
}
