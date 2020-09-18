namespace CafeProject
{
    partial class Register
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Register));
            this.RegNameBox = new Bunifu.Framework.UI.BunifuMaterialTextbox();
            this.RegLastnameBox = new Bunifu.Framework.UI.BunifuMaterialTextbox();
            this.RegUsernameBox = new Bunifu.Framework.UI.BunifuMaterialTextbox();
            this.RegPasswordBox = new Bunifu.Framework.UI.BunifuMaterialTextbox();
            this.RegVerificationBox = new Bunifu.Framework.UI.BunifuMaterialTextbox();
            this.BackBTN = new Bunifu.Framework.UI.BunifuFlatButton();
            this.RegBTN = new Bunifu.Framework.UI.BunifuFlatButton();
            this.RegPositionsComboBox = new Bunifu.Framework.UI.BunifuDropdown();
            this.BackBTN2 = new Bunifu.Framework.UI.BunifuFlatButton();
            this.SuspendLayout();
            // 
            // RegNameBox
            // 
            this.RegNameBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.RegNameBox.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.RegNameBox.Font = new System.Drawing.Font("Century Gothic", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RegNameBox.ForeColor = System.Drawing.Color.White;
            this.RegNameBox.HintForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(176)))), ((int)(((byte)(183)))), ((int)(((byte)(177)))));
            this.RegNameBox.HintText = "სახელი";
            this.RegNameBox.isPassword = false;
            this.RegNameBox.LineFocusedColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(100)))), ((int)(((byte)(15)))));
            this.RegNameBox.LineIdleColor = System.Drawing.Color.Gray;
            this.RegNameBox.LineMouseHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(100)))), ((int)(((byte)(15)))));
            this.RegNameBox.LineThickness = 4;
            this.RegNameBox.Location = new System.Drawing.Point(52, 66);
            this.RegNameBox.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.RegNameBox.Name = "RegNameBox";
            this.RegNameBox.Size = new System.Drawing.Size(268, 47);
            this.RegNameBox.TabIndex = 7;
            this.RegNameBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Left;
            this.RegNameBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.RegNameBox_KeyDown);
            // 
            // RegLastnameBox
            // 
            this.RegLastnameBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.RegLastnameBox.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.RegLastnameBox.Font = new System.Drawing.Font("Century Gothic", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RegLastnameBox.ForeColor = System.Drawing.Color.White;
            this.RegLastnameBox.HintForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(176)))), ((int)(((byte)(183)))), ((int)(((byte)(177)))));
            this.RegLastnameBox.HintText = "გვარი";
            this.RegLastnameBox.isPassword = false;
            this.RegLastnameBox.LineFocusedColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(100)))), ((int)(((byte)(15)))));
            this.RegLastnameBox.LineIdleColor = System.Drawing.Color.Gray;
            this.RegLastnameBox.LineMouseHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(100)))), ((int)(((byte)(15)))));
            this.RegLastnameBox.LineThickness = 4;
            this.RegLastnameBox.Location = new System.Drawing.Point(52, 123);
            this.RegLastnameBox.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.RegLastnameBox.Name = "RegLastnameBox";
            this.RegLastnameBox.Size = new System.Drawing.Size(268, 47);
            this.RegLastnameBox.TabIndex = 8;
            this.RegLastnameBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Left;
            this.RegLastnameBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.RegLastnameBox_KeyDown);
            // 
            // RegUsernameBox
            // 
            this.RegUsernameBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.RegUsernameBox.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.RegUsernameBox.Font = new System.Drawing.Font("Century Gothic", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RegUsernameBox.ForeColor = System.Drawing.Color.White;
            this.RegUsernameBox.HintForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(176)))), ((int)(((byte)(183)))), ((int)(((byte)(177)))));
            this.RegUsernameBox.HintText = "მომხმარებლის სახელი";
            this.RegUsernameBox.isPassword = false;
            this.RegUsernameBox.LineFocusedColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(100)))), ((int)(((byte)(15)))));
            this.RegUsernameBox.LineIdleColor = System.Drawing.Color.Gray;
            this.RegUsernameBox.LineMouseHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(100)))), ((int)(((byte)(15)))));
            this.RegUsernameBox.LineThickness = 4;
            this.RegUsernameBox.Location = new System.Drawing.Point(52, 180);
            this.RegUsernameBox.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.RegUsernameBox.Name = "RegUsernameBox";
            this.RegUsernameBox.Size = new System.Drawing.Size(268, 47);
            this.RegUsernameBox.TabIndex = 9;
            this.RegUsernameBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Left;
            this.RegUsernameBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.RegUsernameBox_KeyDown);
            // 
            // RegPasswordBox
            // 
            this.RegPasswordBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.RegPasswordBox.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.RegPasswordBox.Font = new System.Drawing.Font("Century Gothic", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RegPasswordBox.ForeColor = System.Drawing.Color.White;
            this.RegPasswordBox.HintForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(176)))), ((int)(((byte)(183)))), ((int)(((byte)(177)))));
            this.RegPasswordBox.HintText = "პაროლი";
            this.RegPasswordBox.isPassword = false;
            this.RegPasswordBox.LineFocusedColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(100)))), ((int)(((byte)(15)))));
            this.RegPasswordBox.LineIdleColor = System.Drawing.Color.Gray;
            this.RegPasswordBox.LineMouseHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(100)))), ((int)(((byte)(15)))));
            this.RegPasswordBox.LineThickness = 4;
            this.RegPasswordBox.Location = new System.Drawing.Point(52, 237);
            this.RegPasswordBox.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.RegPasswordBox.Name = "RegPasswordBox";
            this.RegPasswordBox.Size = new System.Drawing.Size(268, 47);
            this.RegPasswordBox.TabIndex = 10;
            this.RegPasswordBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Left;
            this.RegPasswordBox.OnValueChanged += new System.EventHandler(this.RegPasswordBox_OnValueChanged);
            this.RegPasswordBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.RegPasswordBox_KeyDown);
            // 
            // RegVerificationBox
            // 
            this.RegVerificationBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.RegVerificationBox.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.RegVerificationBox.Font = new System.Drawing.Font("Century Gothic", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RegVerificationBox.ForeColor = System.Drawing.Color.White;
            this.RegVerificationBox.HintForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(176)))), ((int)(((byte)(183)))), ((int)(((byte)(177)))));
            this.RegVerificationBox.HintText = "გაიმეორე პაროლი";
            this.RegVerificationBox.isPassword = false;
            this.RegVerificationBox.LineFocusedColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(100)))), ((int)(((byte)(15)))));
            this.RegVerificationBox.LineIdleColor = System.Drawing.Color.Gray;
            this.RegVerificationBox.LineMouseHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(100)))), ((int)(((byte)(15)))));
            this.RegVerificationBox.LineThickness = 4;
            this.RegVerificationBox.Location = new System.Drawing.Point(52, 294);
            this.RegVerificationBox.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.RegVerificationBox.Name = "RegVerificationBox";
            this.RegVerificationBox.Size = new System.Drawing.Size(268, 47);
            this.RegVerificationBox.TabIndex = 11;
            this.RegVerificationBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Left;
            this.RegVerificationBox.OnValueChanged += new System.EventHandler(this.RegVerificationBox_OnValueChanged);
            this.RegVerificationBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.RegVerificationBox_KeyDown);
            // 
            // BackBTN
            // 
            this.BackBTN.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.BackBTN.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.BackBTN.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.BackBTN.BorderRadius = 7;
            this.BackBTN.ButtonText = "";
            this.BackBTN.Cursor = System.Windows.Forms.Cursors.Hand;
            this.BackBTN.DisabledColor = System.Drawing.Color.Gray;
            this.BackBTN.Iconcolor = System.Drawing.Color.Transparent;
            this.BackBTN.Iconimage = ((System.Drawing.Image)(resources.GetObject("BackBTN.Iconimage")));
            this.BackBTN.Iconimage_right = null;
            this.BackBTN.Iconimage_right_Selected = null;
            this.BackBTN.Iconimage_Selected = null;
            this.BackBTN.IconMarginLeft = 0;
            this.BackBTN.IconMarginRight = 0;
            this.BackBTN.IconRightVisible = true;
            this.BackBTN.IconRightZoom = 30D;
            this.BackBTN.IconVisible = true;
            this.BackBTN.IconZoom = 40D;
            this.BackBTN.IsTab = false;
            this.BackBTN.Location = new System.Drawing.Point(3, 3);
            this.BackBTN.Name = "BackBTN";
            this.BackBTN.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.BackBTN.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.BackBTN.OnHoverTextColor = System.Drawing.Color.White;
            this.BackBTN.selected = false;
            this.BackBTN.Size = new System.Drawing.Size(49, 48);
            this.BackBTN.TabIndex = 23;
            this.BackBTN.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.BackBTN.Textcolor = System.Drawing.Color.White;
            this.BackBTN.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BackBTN.Click += new System.EventHandler(this.BackBTN_Click);
            // 
            // RegBTN
            // 
            this.RegBTN.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.RegBTN.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.RegBTN.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.RegBTN.BorderRadius = 0;
            this.RegBTN.ButtonText = "რეგისტრაცია";
            this.RegBTN.Cursor = System.Windows.Forms.Cursors.Hand;
            this.RegBTN.DisabledColor = System.Drawing.Color.Gray;
            this.RegBTN.Iconcolor = System.Drawing.Color.Transparent;
            this.RegBTN.Iconimage = ((System.Drawing.Image)(resources.GetObject("RegBTN.Iconimage")));
            this.RegBTN.Iconimage_right = null;
            this.RegBTN.Iconimage_right_Selected = null;
            this.RegBTN.Iconimage_Selected = null;
            this.RegBTN.IconMarginLeft = 0;
            this.RegBTN.IconMarginRight = 0;
            this.RegBTN.IconRightVisible = true;
            this.RegBTN.IconRightZoom = 0D;
            this.RegBTN.IconVisible = true;
            this.RegBTN.IconZoom = 50D;
            this.RegBTN.IsTab = false;
            this.RegBTN.Location = new System.Drawing.Point(81, 423);
            this.RegBTN.Name = "RegBTN";
            this.RegBTN.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.RegBTN.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.RegBTN.OnHoverTextColor = System.Drawing.Color.White;
            this.RegBTN.selected = false;
            this.RegBTN.Size = new System.Drawing.Size(210, 48);
            this.RegBTN.TabIndex = 24;
            this.RegBTN.Text = "რეგისტრაცია";
            this.RegBTN.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.RegBTN.Textcolor = System.Drawing.Color.White;
            this.RegBTN.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RegBTN.Click += new System.EventHandler(this.RegBTN_Click);
            // 
            // RegPositionsComboBox
            // 
            this.RegPositionsComboBox.BackColor = System.Drawing.Color.Transparent;
            this.RegPositionsComboBox.BorderRadius = 3;
            this.RegPositionsComboBox.ForeColor = System.Drawing.Color.White;
            this.RegPositionsComboBox.Items = new string[0];
            this.RegPositionsComboBox.Location = new System.Drawing.Point(52, 351);
            this.RegPositionsComboBox.Name = "RegPositionsComboBox";
            this.RegPositionsComboBox.NomalColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.RegPositionsComboBox.onHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.RegPositionsComboBox.selectedIndex = -1;
            this.RegPositionsComboBox.Size = new System.Drawing.Size(217, 40);
            this.RegPositionsComboBox.TabIndex = 25;
            this.RegPositionsComboBox.onItemSelected += new System.EventHandler(this.RegPositionsComboBox_onItemSelected);
            // 
            // BackBTN2
            // 
            this.BackBTN2.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.BackBTN2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.BackBTN2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.BackBTN2.BorderRadius = 7;
            this.BackBTN2.ButtonText = "";
            this.BackBTN2.Cursor = System.Windows.Forms.Cursors.Hand;
            this.BackBTN2.DisabledColor = System.Drawing.Color.Gray;
            this.BackBTN2.Iconcolor = System.Drawing.Color.Transparent;
            this.BackBTN2.Iconimage = ((System.Drawing.Image)(resources.GetObject("BackBTN2.Iconimage")));
            this.BackBTN2.Iconimage_right = null;
            this.BackBTN2.Iconimage_right_Selected = null;
            this.BackBTN2.Iconimage_Selected = null;
            this.BackBTN2.IconMarginLeft = 0;
            this.BackBTN2.IconMarginRight = 0;
            this.BackBTN2.IconRightVisible = true;
            this.BackBTN2.IconRightZoom = 30D;
            this.BackBTN2.IconVisible = true;
            this.BackBTN2.IconZoom = 40D;
            this.BackBTN2.IsTab = false;
            this.BackBTN2.Location = new System.Drawing.Point(3, 3);
            this.BackBTN2.Name = "BackBTN2";
            this.BackBTN2.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.BackBTN2.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.BackBTN2.OnHoverTextColor = System.Drawing.Color.White;
            this.BackBTN2.selected = false;
            this.BackBTN2.Size = new System.Drawing.Size(49, 48);
            this.BackBTN2.TabIndex = 26;
            this.BackBTN2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.BackBTN2.Textcolor = System.Drawing.Color.White;
            this.BackBTN2.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BackBTN2.Click += new System.EventHandler(this.BackBTN2_Click);
            // 
            // Register
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.Controls.Add(this.BackBTN2);
            this.Controls.Add(this.RegPositionsComboBox);
            this.Controls.Add(this.RegBTN);
            this.Controls.Add(this.BackBTN);
            this.Controls.Add(this.RegVerificationBox);
            this.Controls.Add(this.RegPasswordBox);
            this.Controls.Add(this.RegUsernameBox);
            this.Controls.Add(this.RegLastnameBox);
            this.Controls.Add(this.RegNameBox);
            this.Name = "Register";
            this.Size = new System.Drawing.Size(372, 491);
            this.Load += new System.EventHandler(this.Register_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private Bunifu.Framework.UI.BunifuMaterialTextbox RegNameBox;
        private Bunifu.Framework.UI.BunifuMaterialTextbox RegLastnameBox;
        private Bunifu.Framework.UI.BunifuMaterialTextbox RegUsernameBox;
        private Bunifu.Framework.UI.BunifuMaterialTextbox RegPasswordBox;
        private Bunifu.Framework.UI.BunifuMaterialTextbox RegVerificationBox;
        private Bunifu.Framework.UI.BunifuFlatButton BackBTN;
        private Bunifu.Framework.UI.BunifuFlatButton RegBTN;
        private Bunifu.Framework.UI.BunifuDropdown RegPositionsComboBox;
        private Bunifu.Framework.UI.BunifuFlatButton BackBTN2;
    }
}
