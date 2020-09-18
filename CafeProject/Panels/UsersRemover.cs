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
    public partial class UsersRemover : UserControl
    {
        public UsersRemover()
        {
            InitializeComponent();
        }

        static public BunifuFlatButton backbutton = new BunifuFlatButton();
        static public BunifuFlatButton backbutton2 = new BunifuFlatButton();
        private void UsersRemover_Load(object sender, EventArgs e)
        {
            backbutton = BackBTN2;
            backbutton2 = BackBTN;
            RemoveBTN = button1;
            Gridviews.basegridview = BaseDataGridView;
        }

        private void BackBTN2_Click(object sender, EventArgs e)
        {
            this.Hide();
            FormsList.forms["RemoveOrCreate"].Show();
            MainWindow.form.Width = 282;
            MainWindow.form.Height = 358;
        }

        static public BunifuFlatButton RemoveBTN = new BunifuFlatButton();
        private void button1_Click(object sender, EventArgs e)
        {
            BindingList<UsersInfo> usersInfos = new BindingList<UsersInfo>();

            foreach (DataGridViewCell item in BaseDataGridView.SelectedCells)
            {
                try
                {
                    if (item.RowIndex > 0)
                    {
                        UsersDatabase.Base.RemoveAt(item.RowIndex);

                        string Filename = "Database.txt";
                        var json = JsonConvert.SerializeObject(UsersDatabase.Base);
                        File.WriteAllText(Filename, json);
                        usersInfos = UsersDatabase.Base;
                        BaseDataGridView.DataSource = usersInfos;
                    }   
                }
                catch (Exception)
                {
     
                }
            }
        }

        private void BackBTN_Click(object sender, EventArgs e)
        {
            this.Hide();
            FormsList.forms["ManagerWindow"].Show();
            MainWindow.form.Width = 282;
            MainWindow.form.Height = 459;
        }
    }
}
