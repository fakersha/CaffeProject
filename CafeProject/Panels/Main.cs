using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.IO;

namespace CafeProject
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        private void Main_Load(object sender, EventArgs e)
        {
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            MainWindow.form = this;
            FormsList.forms.Add("LoadingWindow", loadingWindow);
            FormsList.forms.Add("UserLogin", userLogin);
            FormsList.forms.Add("Register", register);
            FormsList.forms.Add("UsersRemover", usersRemover);
            FormsList.forms.Add("RemoveOrCreate", removeOrCreate);
            FormsList.forms.Add("WaiterWindow", waiterWindow);
            FormsList.forms.Add("AccountantWindow", accountantWindow);
            FormsList.forms.Add("KitchenWindow", kitchenWindow);
            FormsList.forms.Add("ManagerWindow", managerWindow);
            FormsList.forms.Add("ManagerNewProducts", managerNewProducts);
            FormsList.forms.Add("AddImportedProducts", addImportedProducts);


            foreach (var item in FormsList.forms)
            {
                if (item.Key != "LoadingWindow")
                {
                    item.Value.Hide();
                }
            }
        }
    }
}
