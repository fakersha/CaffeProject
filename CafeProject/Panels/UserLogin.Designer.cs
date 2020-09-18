namespace CafeProject
{
    partial class UserLogin
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(UserLogin));
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.backgroundWorker2 = new System.ComponentModel.BackgroundWorker();
            this.UserLoginBTN = new Bunifu.Framework.UI.BunifuFlatButton();
            this.UserLogPasswordBox = new Bunifu.Framework.UI.BunifuMaterialTextbox();
            this.UserLogUsernameBox = new Bunifu.Framework.UI.BunifuMaterialTextbox();
            this.SuspendLayout();
            // 
            // UserLoginBTN
            // 
            this.UserLoginBTN.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.UserLoginBTN.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.UserLoginBTN.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.UserLoginBTN.BorderRadius = 0;
            this.UserLoginBTN.ButtonText = "შესვლა";
            this.UserLoginBTN.Cursor = System.Windows.Forms.Cursors.Hand;
            this.UserLoginBTN.DisabledColor = System.Drawing.Color.Gray;
            this.UserLoginBTN.Iconcolor = System.Drawing.Color.Transparent;
            this.UserLoginBTN.Iconimage = ((System.Drawing.Image)(resources.GetObject("UserLoginBTN.Iconimage")));
            this.UserLoginBTN.Iconimage_right = null;
            this.UserLoginBTN.Iconimage_right_Selected = null;
            this.UserLoginBTN.Iconimage_Selected = null;
            this.UserLoginBTN.IconMarginLeft = 0;
            this.UserLoginBTN.IconMarginRight = 0;
            this.UserLoginBTN.IconRightVisible = true;
            this.UserLoginBTN.IconRightZoom = 0D;
            this.UserLoginBTN.IconVisible = true;
            this.UserLoginBTN.IconZoom = 50D;
            this.UserLoginBTN.IsTab = false;
            this.UserLoginBTN.Location = new System.Drawing.Point(65, 314);
            this.UserLoginBTN.Name = "UserLoginBTN";
            this.UserLoginBTN.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.UserLoginBTN.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.UserLoginBTN.OnHoverTextColor = System.Drawing.Color.White;
            this.UserLoginBTN.selected = false;
            this.UserLoginBTN.Size = new System.Drawing.Size(210, 48);
            this.UserLoginBTN.TabIndex = 2;
            this.UserLoginBTN.Text = "შესვლა";
            this.UserLoginBTN.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.UserLoginBTN.Textcolor = System.Drawing.Color.White;
            this.UserLoginBTN.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.UserLoginBTN.Click += new System.EventHandler(this.UserLoginBTN_Click);
            // 
            // UserLogPasswordBox
            // 
            this.UserLogPasswordBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.UserLogPasswordBox.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.UserLogPasswordBox.Font = new System.Drawing.Font("Century Gothic", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.UserLogPasswordBox.ForeColor = System.Drawing.Color.White;
            this.UserLogPasswordBox.HintForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(176)))), ((int)(((byte)(183)))), ((int)(((byte)(177)))));
            this.UserLogPasswordBox.HintText = "პაროლი";
            this.UserLogPasswordBox.isPassword = false;
            this.UserLogPasswordBox.LineFocusedColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(100)))), ((int)(((byte)(15)))));
            this.UserLogPasswordBox.LineIdleColor = System.Drawing.Color.Gray;
            this.UserLogPasswordBox.LineMouseHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(100)))), ((int)(((byte)(15)))));
            this.UserLogPasswordBox.LineThickness = 4;
            this.UserLogPasswordBox.Location = new System.Drawing.Point(33, 135);
            this.UserLogPasswordBox.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.UserLogPasswordBox.Name = "UserLogPasswordBox";
            this.UserLogPasswordBox.Size = new System.Drawing.Size(268, 47);
            this.UserLogPasswordBox.TabIndex = 5;
            this.UserLogPasswordBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Left;
            this.UserLogPasswordBox.OnValueChanged += new System.EventHandler(this.UserLogPasswordBox_OnValueChanged);
            this.UserLogPasswordBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.UserLogPasswordBox_KeyDown);
            // 
            // UserLogUsernameBox
            // 
            this.UserLogUsernameBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.UserLogUsernameBox.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.UserLogUsernameBox.Font = new System.Drawing.Font("Century Gothic", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.UserLogUsernameBox.ForeColor = System.Drawing.Color.White;
            this.UserLogUsernameBox.HintForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(176)))), ((int)(((byte)(183)))), ((int)(((byte)(177)))));
            this.UserLogUsernameBox.HintText = "მომხმარებლის სახელი";
            this.UserLogUsernameBox.isPassword = false;
            this.UserLogUsernameBox.LineFocusedColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(100)))), ((int)(((byte)(15)))));
            this.UserLogUsernameBox.LineIdleColor = System.Drawing.Color.Gray;
            this.UserLogUsernameBox.LineMouseHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(100)))), ((int)(((byte)(15)))));
            this.UserLogUsernameBox.LineThickness = 4;
            this.UserLogUsernameBox.Location = new System.Drawing.Point(33, 64);
            this.UserLogUsernameBox.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.UserLogUsernameBox.Name = "UserLogUsernameBox";
            this.UserLogUsernameBox.Size = new System.Drawing.Size(268, 47);
            this.UserLogUsernameBox.TabIndex = 6;
            this.UserLogUsernameBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Left;
            this.UserLogUsernameBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.UserLogUsernameBox_KeyDown);
            // 
            // UserLogin
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.Controls.Add(this.UserLogUsernameBox);
            this.Controls.Add(this.UserLogPasswordBox);
            this.Controls.Add(this.UserLoginBTN);
            this.ForeColor = System.Drawing.Color.White;
            this.Name = "UserLogin";
            this.Size = new System.Drawing.Size(338, 434);
            this.Load += new System.EventHandler(this.UserLogin_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.ComponentModel.BackgroundWorker backgroundWorker2;
        private Bunifu.Framework.UI.BunifuFlatButton UserLoginBTN;
        private Bunifu.Framework.UI.BunifuMaterialTextbox UserLogPasswordBox;
        private Bunifu.Framework.UI.BunifuMaterialTextbox UserLogUsernameBox;
    }
}
