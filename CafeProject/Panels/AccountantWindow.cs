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
    public partial class AccountantWindow : UserControl
    {
        public AccountantWindow()
        {
            InitializeComponent();
        }

        CMessageBox cMessageBox = new CMessageBox();
        private void AccountantWindow_Load(object sender, EventArgs e)
        {
            Gridviews.accountantgridview = Accountantgridview;
            foreach (var item in Enum.GetValues(typeof(ProductType)))
            {
                FilterComboBox.AddItem(item.ToString());
            }

            Accountantgridview.DataSource = Accountant.AccountantList;
            Accountantgridview.Refresh();
        }

        private void LogoutBTN_Click(object sender, EventArgs e)
        {
            this.Hide();
            MainWindow.form.Width = 353;
            MainWindow.form.Height = 473;
            FormsList.forms["UserLogin"].Show();
        }

        private void ClearBTN_Click(object sender, EventArgs e)
        {
            Accountant.AccountantList.Clear();
            string Filename = "AccountantProductBase.txt";
            var json = JsonConvert.SerializeObject(Accountant.AccountantList);
            File.WriteAllText(Filename, json);
            Accountant.CopyAccountantList.Clear();
            Accountant.PayCopyAccountantList.Clear();
        }

        private void AccountmentSum_Click(object sender, EventArgs e)
        {
            double Sum = new double();
            if (Accountantgridview.DataSource == Accountant.AccountantList)
            {
                foreach (var item in Accountant.AccountantList)
                {
                    Sum += item.Price;

                }
                cMessageBox.Show($"ანგარიში: {Sum}₾");
                Sum = 0;
            }
            else if (Accountantgridview.DataSource == Accountant.CopyAccountantList)
            {
                foreach (var item in Accountant.CopyAccountantList)
                {
                    Sum += item.Price;

                }
                cMessageBox.Show($"ანგარიში: {Sum}₾");
                Sum = 0;
            }
            else if (Accountantgridview.DataSource == Accountant.PayCopyAccountantList)
            {
                foreach (var item in Accountant.PayCopyAccountantList)
                {
                    Sum += item.Price;

                }
                cMessageBox.Show($"ანგარიში: {Sum}₾");
                Sum = 0;
            }
        }

        static public int plus = 0;
        private void BrowseBTN_Click(object sender, EventArgs e)
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
            worksheet.Name = "AccountantList";

            for (int i = 1; i < Accountantgridview.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = Accountantgridview.Columns[i - 1].HeaderText;
            }

            for (int i = 0; i < Accountantgridview.Rows.Count - 1; i++)
            {
                for (int j = 0; j < Accountantgridview.Columns.Count; j++)
                {
                   
                        worksheet.Cells[i + 2, j + 1] = Accountantgridview.Rows[i].Cells[j].Value.ToString();
                }
            }

            plus++;

            FileStream sb = new FileStream("PlusNumber.txt", FileMode.OpenOrCreate);
            StreamWriter sw = new StreamWriter(sb);
            sw.Write(plus);
            sw.Close();

            using (StreamReader st = new StreamReader("PlusNumber.txt"))
            {
                string text = st.ReadToEnd();
                var result = int.Parse(text);
                plus = result;
            }

            string excelname = (savetext + $"/AccountantList{plus}.xlsx");
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

        private void FilterComboBox_onItemSelected(object sender, EventArgs e)
        {
            Accountant.CopyAccountantList = new BindingList<OrderItem>();
            foreach (var item in Accountant.AccountantList)
            {
                if (FilterComboBox.selectedValue.ToString() == item.Type.ToString())
                {
                    Accountant.CopyAccountantList.Add(item);
                }
            }

            if (FilterComboBox.selectedValue.ToString() == ProductType.მენიუ.ToString())
            {
                Accountantgridview.DataSource = Accountant.AccountantList;
            }
            else
            {
                Accountantgridview.DataSource = Accountant.CopyAccountantList;
            }
            Accountantgridview.Refresh();
        }

    }
}
