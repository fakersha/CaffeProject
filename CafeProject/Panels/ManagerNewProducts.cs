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
    public partial class ManagerNewProducts : UserControl
    {
        public ManagerNewProducts()
        {
            InitializeComponent();
        }

        CMessageBox cMessageBox = new CMessageBox();
        static public BunifuFlatButton bottomBTN1 = new BunifuFlatButton();
        static public BunifuFlatButton bottomBTN2 = new BunifuFlatButton();
        private void ManagerNewProducts_Load(object sender, EventArgs e)
        {
            bottomBTN1 = bunifuFlatButton1;
            bottomBTN2 = bunifuFlatButton2;
            foreach (var item in Enum.GetValues(typeof(ProductType)))
            {
                if (item.ToString() != ProductType.მენიუ.ToString())
                {
                    TypeBox.AddItem(item.ToString());
                }
            }

            TypeBox.selectedIndex = 0;
            MenuGridView.DataSource = Menu.MenuItems;
        }

        private void BackBTN_Click(object sender, EventArgs e)
        {
            this.Hide();
            FormsList.forms["ManagerWindow"].Show();
            MainWindow.form.Width = 282;
            MainWindow.form.Height = 459;
        }

        private void AddBTN_Click(object sender, EventArgs e)
        {
            string name = NameBox?.Text;
            string type = TypeBox.selectedValue?.ToString();
            int ID = new int();
            double price = new double();
            try
            {
                price = double.Parse(PriceBox.Text);
                ID = int.Parse(IDBox.Text);
            }
            catch (Exception)
            {
                cMessageBox.Show("პროდუქტი ვერ დაემატა!");
            }
           
            
            bool yesorno = false;
            foreach (var item in Menu.MenuItems)
            {
                if (name == item.Name)
                {
                    yesorno = true;
                }

            }

            if (yesorno == false)
            {
                if (name != null && price != 0 && type != null && ID != 0)
                {
                    Menu.AddInMenu(new MenuItem(NameBox.Text, price, type, ID));
                    cMessageBox.Show("პროდუქტი დაემატა!");
                    MenuGridView.Refresh();
                    NameBox.Text = "";
                    PriceBox.Text = "";
                    IDBox.Text = "";
                    TypeBox.Text = "";
                }
            }
            else
            {
                cMessageBox.Show("მსგავსი პროდუქტი დამატებულია!");
            }
        }

        private void bunifuFlatButton1_Click(object sender, EventArgs e)
        {
            
            MainWindow.form.Width = 463;
            MainWindow.form.Height = 689;
            MenuGridView.Refresh();
            bunifuFlatButton2.Show();
            bunifuFlatButton1.Hide();

        }

        private void bunifuFlatButton2_Click(object sender, EventArgs e)
        {
            
            MainWindow.form.Width = 463;
            MainWindow.form.Height = 480;
            bunifuFlatButton1.Show();
            bunifuFlatButton2.Hide();
        }

        private void RemoveClick_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell item in MenuGridView.SelectedCells)
            {
                try
                {
                    Menu.MenuItems.RemoveAt(item.RowIndex);
                }
                catch (Exception)
                {

                }
            }
        }

        static public int plus = new int();
        private void SaveBTN_Click(object sender, EventArgs e)
        {
            string savetext = "";
            using (FolderBrowserDialog fbd = new FolderBrowserDialog() { Description = "Select your path" })
            {
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    savetext = fbd.SelectedPath;
                }
            }

            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            app.Visible = false;
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            worksheet.Name = "Menu";

            for (int i = 1; i < MenuGridView.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = MenuGridView.Columns[i - 1].HeaderText;
            }

            for (int i = 0; i < MenuGridView.Rows.Count - 1; i++)
            {
                for (int j = 0; j < MenuGridView.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = MenuGridView.Rows[i].Cells[j].Value.ToString();
                }
            }

            plus++;

            FileStream sb = new FileStream("MenuPlusNumber.txt", FileMode.OpenOrCreate);
            StreamWriter sw = new StreamWriter(sb);
            sw.Write(plus);
            sw.Close();

            using (StreamReader st = new StreamReader("MenuPlusNumber.txt"))
            {
                string text = st.ReadToEnd();
                var result = int.Parse(text);
                plus = result;
            }

            string excelname = (savetext + $"/Menu{plus}.xlsx");
            try
            {
                workbook.SaveAs(excelname, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            catch (Exception)
            {
                cMessageBox.Show("ეს ფაილი უკვე არსებობს!");
            }

            workbook.Close();
        }

        private void NameBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                AddBTN_Click(null, null);
            }
        }

        private void PriceBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                AddBTN_Click(null, null);
            }
        }

        private void IDBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                AddBTN_Click(null, null);
            }
        }
    }
}
