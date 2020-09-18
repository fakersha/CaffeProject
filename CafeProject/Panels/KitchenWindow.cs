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
    public partial class KitchenWindow : UserControl
    {
        public KitchenWindow()
        {
            InitializeComponent();
            KitchenTables.tables = new List<BunifuFlatButton>()
            {
                Table_1, Table_2, Table_3, Table_4, Table_5, Table_6, Table_7, Table_8
            };
        }

        CMessageBox cMessageBox = new CMessageBox();

        private void BTNColors()
        {
            if (KitchenTables.tables[KitchenTables.Which].Normalcolor != Color.FromArgb(247, 227, 96, 40))
            {
                KitchenTables.tables[KitchenTables.Which].Normalcolor = Color.FromArgb(220, 151, 60, 49);
            }
            foreach (var item in KitchenTables.tables)
            {
                if (item != KitchenTables.tables[KitchenTables.Which] && item.Normalcolor != Color.FromArgb(247, 227, 96, 40))
                {
                    item.Normalcolor = Color.FromArgb(220, 151, 60, 49);
                }
            }
        }

        private void Description()
        {
            if (Waitress.Descriptions[KitchenTables.Which] != "")
            {
                cMessageBox.Show(Waitress.Descriptions[KitchenTables.Which]);
            }
        }

        

        private void KitchenWindow_Load(object sender, EventArgs e)
        {

        }

        private void LogoutBTN_Click(object sender, EventArgs e)
        {
            this.Hide();
            MainWindow.form.Width = 353;
            MainWindow.form.Height = 473;
            FormsList.forms["UserLogin"].Show();
        }

        private void Table_1_Click(object sender, EventArgs e)
        {
            KitchenGridView.DataSource = KitchenTables.TableList[0];
            KitchenTables.Which = 0;
            Description();
            BTNColors();
        }

        private void Table_2_Click(object sender, EventArgs e)
        {
            KitchenGridView.DataSource = KitchenTables.TableList[1];
            KitchenTables.Which = 1;
            Description();
            BTNColors();
        }

        private void Table_3_Click(object sender, EventArgs e)
        {
            KitchenGridView.DataSource = KitchenTables.TableList[2];
            KitchenTables.Which = 2;
            Description();
            BTNColors();
        }

        private void DoneBTN_Click(object sender, EventArgs e)
        {
            int tableindex = KitchenTables.Which + 1;
            ResultBox.resultcombobox.ForeColor = Color.Red;
            ResultBox.resultcombobox.AddItem(($"მაგიდა {tableindex} მზადაა!"));
        }

        private void ClearBTN_Click(object sender, EventArgs e)
        {
            Waitress.Descriptions[KitchenTables.Which] = "";
            KitchenTables.TableList[KitchenTables.Which].Clear();
            KitchenTables.tables[KitchenTables.Which].Normalcolor = Color.FromArgb(220, 151, 60, 49);
        }

        private void Table_5_Click(object sender, EventArgs e)
        {
            KitchenGridView.DataSource = KitchenTables.TableList[4];
            KitchenTables.Which = 4;
            Description();
            BTNColors();
        }

        private void Table_6_Click(object sender, EventArgs e)
        {
            KitchenGridView.DataSource = KitchenTables.TableList[5];
            KitchenTables.Which = 5;
            Description();
            BTNColors();
        }

        private void Table_7_Click(object sender, EventArgs e)
        {
            KitchenGridView.DataSource = KitchenTables.TableList[6];
            KitchenTables.Which = 6;
            Description();
            BTNColors();
        }

        private void Table_8_Click(object sender, EventArgs e)
        {
            KitchenGridView.DataSource = KitchenTables.TableList[7];
            KitchenTables.Which = 7;
            Description();
            BTNColors();
        }

        private void Table_4_Click(object sender, EventArgs e)
        {
            KitchenGridView.DataSource = KitchenTables.TableList[3];
            KitchenTables.Which = 3;
            Description();
            BTNColors();
        }

    }
}
