using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CafeProject
{
    public partial class LoadingWindow : UserControl
    {
        public LoadingWindow()
        {
            InitializeComponent();
        }

        private void bunifuGradientPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void LoadingWindow_Load(object sender, EventArgs e)
        {
            LoadingTimer.Start();
            SecondTimer.Start();
                
        }

        int a = 10;
        int i = 0;
        private void LoadingTimer_Tick(object sender, EventArgs e)
        {
            a += 30;
            LoadingBar.Value = a;
            if (LoadingBar.Value == 40)
            {
                LoadingImage3.Hide();
                Procents.Text = "40%";
            }
            else if (LoadingBar.Value == 70)
            {
                LoadingImage2.Hide();
                Procents.Text = "70%";
            }
            if (LoadingBar.Value == 100)
            {
                Procents.Text = "100%";
                LoadingImage3.Show();
                LoadingTimer.Stop();

            }
        }

        private void SecondTimer_Tick(object sender, EventArgs e)
        {
            i++;
            if (i == 0)
            {
                LoadingPoint.Text = ".";
            }
            else if (i == 1)
            {
                LoadingPoint.Text = "..";
            }
            else if (i == 2)
            {
                LoadingPoint.Text = "...";
            }
            else if (i == 3)
            {
                LoadingPoint.Text = ".";
            }
            else if (i == 4)
            {
                LoadingPoint.Text = "..";
            }
            else if (i == 5)
            {
                LoadingPoint.Text = "...";
            }
            else if (i == 6)
            {
                LoadingPoint.Text = ".";
                SecondTimer.Stop();
                this.Hide();
                MainWindow.form.Width = 353;
                MainWindow.form.Height = 473;
                FormsList.forms["UserLogin"].Show();
            }
        }
    }
}
