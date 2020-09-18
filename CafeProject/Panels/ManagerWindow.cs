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
    public partial class ManagerWindow : UserControl
    {
        public ManagerWindow()
        {
            InitializeComponent();
        }

        CMessageBox cMessageBox = new CMessageBox();
        private void ManagerWindow_Load(object sender, EventArgs e)
        {

        }

        private void RegisterBTN_Click(object sender, EventArgs e)
        {
            this.Hide();
            MainWindow.form.Width = 387;
            MainWindow.form.Height = 530;
            FormsList.forms["Register"].Show();
        }

        private void UserRemoveBTN_Click(object sender, EventArgs e)
        {
            this.Hide();
            MainWindow.form.Width = 610;
            MainWindow.form.Height = 369;
            FormsList.forms["UsersRemover"].Show();
            string Filename = "Database.txt";
            var json = JsonConvert.SerializeObject(UsersDatabase.Base);
            File.WriteAllText(Filename, json);

            using (StreamReader st = new StreamReader(Filename))
            {
                string text = st.ReadToEnd();
                var result = JsonConvert.DeserializeObject<BindingList<UsersInfo>>(text);
                UsersDatabase.Base = result;
                Gridviews.basegridview.DataSource = UsersDatabase.Base;
                Gridviews.basegridview.Width = 441;
                Gridviews.basegridview.Height = 163;
                Gridviews.basegridview.Left = 77;
                UsersRemover.RemoveBTN.Left = 77;
                Gridviews.basegridview.Columns[3].Visible = false;
            }
        }

        private void AddProductBTN_Click(object sender, EventArgs e)
        {
            this.Hide();
            FormsList.forms["ManagerNewProducts"].Show();
            MainWindow.form.Width = 463;
            MainWindow.form.Height = 480;
            ManagerNewProducts.bottomBTN2.Hide();
            ManagerNewProducts.bottomBTN1.Show();
        }

        private void AddNewProducts_Click(object sender, EventArgs e)
        {
            this.Hide();
            FormsList.forms["AddImportedProducts"].Show();
            MainWindow.form.Width = 463;
            MainWindow.form.Height = 480;
            AddImportedProducts.bottomBTN2.Hide();
            AddImportedProducts.bottomBTN1.Show();
        }

        private void LogoutBTN_Click(object sender, EventArgs e)
        {
            this.Hide();
            MainWindow.form.Width = 353;
            MainWindow.form.Height = 473;
            FormsList.forms["UserLogin"].Show();
            
        }
    }
}
