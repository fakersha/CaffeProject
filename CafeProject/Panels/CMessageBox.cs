using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CafeProject
{
    public partial class CMessageBox : Form
    {
        public CMessageBox()
        {
            InitializeComponent();
        }

        private void ConfirmBTN_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        public void Show(string text)
        {
            CMessageBox cMessageBox = new CMessageBox();
            cMessageBox.MessageBoxText.Text = text;
            int size = (cMessageBox.Width - cMessageBox.MessageBoxText.Width) / 2;
            cMessageBox.MessageBoxText.Left = size;
            cMessageBox.Show();
        }
    }
}
