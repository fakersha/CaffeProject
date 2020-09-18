namespace CafeProject
{
    partial class Main
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Main));
            this.userLogin = new CafeProject.UserLogin();
            this.removeOrCreate = new CafeProject.RemoveOrCreate();
            this.register = new CafeProject.Register();
            this.usersRemover = new CafeProject.UsersRemover();
            this.waiterWindow = new CafeProject.WaiterWindow();
            this.accountantWindow = new CafeProject.AccountantWindow();
            this.kitchenWindow = new CafeProject.KitchenWindow();
            this.managerWindow = new CafeProject.ManagerWindow();
            this.managerNewProducts = new CafeProject.ManagerNewProducts();
            this.addImportedProducts = new CafeProject.AddImportedProducts();
            this.loadingWindow = new CafeProject.LoadingWindow();
            this.SuspendLayout();
            // 
            // userLogin
            // 
            this.userLogin.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("userLogin.BackgroundImage")));
            this.userLogin.ForeColor = System.Drawing.Color.White;
            this.userLogin.Location = new System.Drawing.Point(0, 0);
            this.userLogin.Name = "userLogin";
            this.userLogin.Size = new System.Drawing.Size(338, 434);
            this.userLogin.TabIndex = 0;
            // 
            // removeOrCreate
            // 
            this.removeOrCreate.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("removeOrCreate.BackgroundImage")));
            this.removeOrCreate.Location = new System.Drawing.Point(0, 0);
            this.removeOrCreate.Name = "removeOrCreate";
            this.removeOrCreate.Size = new System.Drawing.Size(267, 319);
            this.removeOrCreate.TabIndex = 1;
            // 
            // register
            // 
            this.register.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("register.BackgroundImage")));
            this.register.Location = new System.Drawing.Point(0, 0);
            this.register.Name = "register";
            this.register.Size = new System.Drawing.Size(372, 491);
            this.register.TabIndex = 2;
            // 
            // usersRemover
            // 
            this.usersRemover.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("usersRemover.BackgroundImage")));
            this.usersRemover.Location = new System.Drawing.Point(0, 0);
            this.usersRemover.Name = "usersRemover";
            this.usersRemover.Size = new System.Drawing.Size(595, 330);
            this.usersRemover.TabIndex = 3;
            // 
            // waiterWindow
            // 
            this.waiterWindow.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("waiterWindow.BackgroundImage")));
            this.waiterWindow.Location = new System.Drawing.Point(0, 0);
            this.waiterWindow.Name = "waiterWindow";
            this.waiterWindow.Size = new System.Drawing.Size(1192, 494);
            this.waiterWindow.TabIndex = 4;
            // 
            // accountantWindow
            // 
            this.accountantWindow.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("accountantWindow.BackgroundImage")));
            this.accountantWindow.Location = new System.Drawing.Point(0, 0);
            this.accountantWindow.Name = "accountantWindow";
            this.accountantWindow.Size = new System.Drawing.Size(606, 533);
            this.accountantWindow.TabIndex = 5;
            // 
            // kitchenWindow
            // 
            this.kitchenWindow.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("kitchenWindow.BackgroundImage")));
            this.kitchenWindow.Location = new System.Drawing.Point(0, 0);
            this.kitchenWindow.Name = "kitchenWindow";
            this.kitchenWindow.Size = new System.Drawing.Size(589, 505);
            this.kitchenWindow.TabIndex = 6;
            // 
            // managerWindow
            // 
            this.managerWindow.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("managerWindow.BackgroundImage")));
            this.managerWindow.Location = new System.Drawing.Point(0, 0);
            this.managerWindow.Name = "managerWindow";
            this.managerWindow.Size = new System.Drawing.Size(267, 420);
            this.managerWindow.TabIndex = 7;
            // 
            // managerNewProducts
            // 
            this.managerNewProducts.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("managerNewProducts.BackgroundImage")));
            this.managerNewProducts.Location = new System.Drawing.Point(0, 0);
            this.managerNewProducts.Name = "managerNewProducts";
            this.managerNewProducts.Size = new System.Drawing.Size(448, 650);
            this.managerNewProducts.TabIndex = 8;
            // 
            // addImportedProducts
            // 
            this.addImportedProducts.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("addImportedProducts.BackgroundImage")));
            this.addImportedProducts.Location = new System.Drawing.Point(0, 0);
            this.addImportedProducts.Name = "addImportedProducts";
            this.addImportedProducts.Size = new System.Drawing.Size(448, 650);
            this.addImportedProducts.TabIndex = 9;
            // 
            // loadingWindow
            // 
            this.loadingWindow.Location = new System.Drawing.Point(0, 0);
            this.loadingWindow.Name = "loadingWindow";
            this.loadingWindow.Size = new System.Drawing.Size(676, 476);
            this.loadingWindow.TabIndex = 10;
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(676, 476);
            this.Controls.Add(this.loadingWindow);
            this.Controls.Add(this.addImportedProducts);
            this.Controls.Add(this.managerNewProducts);
            this.Controls.Add(this.managerWindow);
            this.Controls.Add(this.kitchenWindow);
            this.Controls.Add(this.accountantWindow);
            this.Controls.Add(this.waiterWindow);
            this.Controls.Add(this.usersRemover);
            this.Controls.Add(this.register);
            this.Controls.Add(this.removeOrCreate);
            this.Controls.Add(this.userLogin);
            this.Name = "Main";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Main_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private UserLogin userLogin;
        private RemoveOrCreate removeOrCreate;
        private Register register;
        private UsersRemover usersRemover;
        private WaiterWindow waiterWindow;
        private AccountantWindow accountantWindow;
        private KitchenWindow kitchenWindow;
        private ManagerWindow managerWindow;
        private ManagerNewProducts managerNewProducts;
        private AddImportedProducts addImportedProducts;
        private LoadingWindow loadingWindow;
    }
}

