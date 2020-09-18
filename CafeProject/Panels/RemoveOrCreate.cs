﻿using System;
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
    public partial class RemoveOrCreate : UserControl
    {
        public RemoveOrCreate()
        {
            InitializeComponent();
        }

        private void RemoveOrCreate_Load(object sender, EventArgs e)
        {

        }

        private void LogOutBTN_Click(object sender, EventArgs e)
        {
            this.Hide();
            MainWindow.form.Width = 353;
            MainWindow.form.Height = 473;
            FormsList.forms["UserLogin"].Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            MainWindow.form.Width = 387;
            MainWindow.form.Height = 530;
            FormsList.forms["Register"].Show();
        }

        private void button2_Click(object sender, EventArgs e)
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
                Gridviews.basegridview.Width = 541;
                Gridviews.basegridview.Height = 163;
                Gridviews.basegridview.Left = 22;
                UsersRemover.RemoveBTN.Left = 22;
                Gridviews.basegridview.Columns[3].Visible = true;
            }
        }
    }
}
