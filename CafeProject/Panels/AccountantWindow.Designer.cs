namespace CafeProject
{
    partial class AccountantWindow
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AccountantWindow));
            this.Accountantgridview = new Bunifu.Framework.UI.BunifuCustomDataGrid();
            this.LogoutBTN = new Bunifu.Framework.UI.BunifuFlatButton();
            this.ClearBTN = new Bunifu.Framework.UI.BunifuFlatButton();
            this.AccountmentSum = new Bunifu.Framework.UI.BunifuFlatButton();
            this.BrowseBTN = new Bunifu.Framework.UI.BunifuFlatButton();
            this.FilterComboBox = new Bunifu.Framework.UI.BunifuDropdown();
            ((System.ComponentModel.ISupportInitialize)(this.Accountantgridview)).BeginInit();
            this.SuspendLayout();
            // 
            // Accountantgridview
            // 
            this.Accountantgridview.AllowUserToDeleteRows = false;
            this.Accountantgridview.AllowUserToResizeColumns = false;
            this.Accountantgridview.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.Accountantgridview.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.Accountantgridview.BackgroundColor = System.Drawing.Color.Gainsboro;
            this.Accountantgridview.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.Accountantgridview.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Century Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.Accountantgridview.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.Accountantgridview.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.Accountantgridview.DoubleBuffered = true;
            this.Accountantgridview.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.Accountantgridview.EnableHeadersVisualStyles = false;
            this.Accountantgridview.HeaderBgColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.Accountantgridview.HeaderForeColor = System.Drawing.Color.White;
            this.Accountantgridview.Location = new System.Drawing.Point(33, 65);
            this.Accountantgridview.Name = "Accountantgridview";
            this.Accountantgridview.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            this.Accountantgridview.Size = new System.Drawing.Size(541, 255);
            this.Accountantgridview.TabIndex = 2;
            // 
            // LogoutBTN
            // 
            this.LogoutBTN.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.LogoutBTN.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.LogoutBTN.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.LogoutBTN.BorderRadius = 7;
            this.LogoutBTN.ButtonText = "";
            this.LogoutBTN.Cursor = System.Windows.Forms.Cursors.Hand;
            this.LogoutBTN.DisabledColor = System.Drawing.Color.Gray;
            this.LogoutBTN.Iconcolor = System.Drawing.Color.Transparent;
            this.LogoutBTN.Iconimage = ((System.Drawing.Image)(resources.GetObject("LogoutBTN.Iconimage")));
            this.LogoutBTN.Iconimage_right = null;
            this.LogoutBTN.Iconimage_right_Selected = null;
            this.LogoutBTN.Iconimage_Selected = null;
            this.LogoutBTN.IconMarginLeft = 0;
            this.LogoutBTN.IconMarginRight = 0;
            this.LogoutBTN.IconRightVisible = true;
            this.LogoutBTN.IconRightZoom = 30D;
            this.LogoutBTN.IconVisible = true;
            this.LogoutBTN.IconZoom = 40D;
            this.LogoutBTN.IsTab = false;
            this.LogoutBTN.Location = new System.Drawing.Point(3, 3);
            this.LogoutBTN.Name = "LogoutBTN";
            this.LogoutBTN.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.LogoutBTN.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.LogoutBTN.OnHoverTextColor = System.Drawing.Color.White;
            this.LogoutBTN.selected = false;
            this.LogoutBTN.Size = new System.Drawing.Size(49, 48);
            this.LogoutBTN.TabIndex = 23;
            this.LogoutBTN.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.LogoutBTN.Textcolor = System.Drawing.Color.White;
            this.LogoutBTN.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LogoutBTN.Click += new System.EventHandler(this.LogoutBTN_Click);
            // 
            // ClearBTN
            // 
            this.ClearBTN.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.ClearBTN.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.ClearBTN.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClearBTN.BorderRadius = 0;
            this.ClearBTN.ButtonText = "გასუფთავება";
            this.ClearBTN.Cursor = System.Windows.Forms.Cursors.Hand;
            this.ClearBTN.DisabledColor = System.Drawing.Color.Gray;
            this.ClearBTN.Iconcolor = System.Drawing.Color.Transparent;
            this.ClearBTN.Iconimage = ((System.Drawing.Image)(resources.GetObject("ClearBTN.Iconimage")));
            this.ClearBTN.Iconimage_right = null;
            this.ClearBTN.Iconimage_right_Selected = null;
            this.ClearBTN.Iconimage_Selected = null;
            this.ClearBTN.IconMarginLeft = 0;
            this.ClearBTN.IconMarginRight = 0;
            this.ClearBTN.IconRightVisible = true;
            this.ClearBTN.IconRightZoom = 0D;
            this.ClearBTN.IconVisible = true;
            this.ClearBTN.IconZoom = 50D;
            this.ClearBTN.IsTab = false;
            this.ClearBTN.Location = new System.Drawing.Point(-2, 458);
            this.ClearBTN.Name = "ClearBTN";
            this.ClearBTN.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.ClearBTN.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.ClearBTN.OnHoverTextColor = System.Drawing.Color.White;
            this.ClearBTN.selected = false;
            this.ClearBTN.Size = new System.Drawing.Size(203, 76);
            this.ClearBTN.TabIndex = 24;
            this.ClearBTN.Text = "გასუფთავება";
            this.ClearBTN.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.ClearBTN.Textcolor = System.Drawing.Color.White;
            this.ClearBTN.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ClearBTN.Click += new System.EventHandler(this.ClearBTN_Click);
            // 
            // AccountmentSum
            // 
            this.AccountmentSum.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.AccountmentSum.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.AccountmentSum.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.AccountmentSum.BorderRadius = 0;
            this.AccountmentSum.ButtonText = "ანგარიში";
            this.AccountmentSum.Cursor = System.Windows.Forms.Cursors.Hand;
            this.AccountmentSum.DisabledColor = System.Drawing.Color.Gray;
            this.AccountmentSum.Iconcolor = System.Drawing.Color.Transparent;
            this.AccountmentSum.Iconimage = ((System.Drawing.Image)(resources.GetObject("AccountmentSum.Iconimage")));
            this.AccountmentSum.Iconimage_right = null;
            this.AccountmentSum.Iconimage_right_Selected = null;
            this.AccountmentSum.Iconimage_Selected = null;
            this.AccountmentSum.IconMarginLeft = 0;
            this.AccountmentSum.IconMarginRight = 0;
            this.AccountmentSum.IconRightVisible = true;
            this.AccountmentSum.IconRightZoom = 0D;
            this.AccountmentSum.IconVisible = true;
            this.AccountmentSum.IconZoom = 50D;
            this.AccountmentSum.IsTab = false;
            this.AccountmentSum.Location = new System.Drawing.Point(201, 458);
            this.AccountmentSum.Name = "AccountmentSum";
            this.AccountmentSum.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.AccountmentSum.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.AccountmentSum.OnHoverTextColor = System.Drawing.Color.White;
            this.AccountmentSum.selected = false;
            this.AccountmentSum.Size = new System.Drawing.Size(203, 76);
            this.AccountmentSum.TabIndex = 26;
            this.AccountmentSum.Text = "ანგარიში";
            this.AccountmentSum.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.AccountmentSum.Textcolor = System.Drawing.Color.White;
            this.AccountmentSum.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.AccountmentSum.Click += new System.EventHandler(this.AccountmentSum_Click);
            // 
            // BrowseBTN
            // 
            this.BrowseBTN.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.BrowseBTN.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.BrowseBTN.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.BrowseBTN.BorderRadius = 0;
            this.BrowseBTN.ButtonText = "შენახვა";
            this.BrowseBTN.Cursor = System.Windows.Forms.Cursors.Hand;
            this.BrowseBTN.DisabledColor = System.Drawing.Color.Gray;
            this.BrowseBTN.Iconcolor = System.Drawing.Color.Transparent;
            this.BrowseBTN.Iconimage = ((System.Drawing.Image)(resources.GetObject("BrowseBTN.Iconimage")));
            this.BrowseBTN.Iconimage_right = null;
            this.BrowseBTN.Iconimage_right_Selected = null;
            this.BrowseBTN.Iconimage_Selected = null;
            this.BrowseBTN.IconMarginLeft = 0;
            this.BrowseBTN.IconMarginRight = 0;
            this.BrowseBTN.IconRightVisible = true;
            this.BrowseBTN.IconRightZoom = 0D;
            this.BrowseBTN.IconVisible = true;
            this.BrowseBTN.IconZoom = 50D;
            this.BrowseBTN.IsTab = false;
            this.BrowseBTN.Location = new System.Drawing.Point(404, 458);
            this.BrowseBTN.Name = "BrowseBTN";
            this.BrowseBTN.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.BrowseBTN.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.BrowseBTN.OnHoverTextColor = System.Drawing.Color.White;
            this.BrowseBTN.selected = false;
            this.BrowseBTN.Size = new System.Drawing.Size(203, 76);
            this.BrowseBTN.TabIndex = 27;
            this.BrowseBTN.Text = "შენახვა";
            this.BrowseBTN.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.BrowseBTN.Textcolor = System.Drawing.Color.White;
            this.BrowseBTN.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BrowseBTN.Click += new System.EventHandler(this.BrowseBTN_Click);
            // 
            // FilterComboBox
            // 
            this.FilterComboBox.BackColor = System.Drawing.Color.Transparent;
            this.FilterComboBox.BorderRadius = 3;
            this.FilterComboBox.ForeColor = System.Drawing.Color.White;
            this.FilterComboBox.Items = new string[0];
            this.FilterComboBox.Location = new System.Drawing.Point(33, 326);
            this.FilterComboBox.Name = "FilterComboBox";
            this.FilterComboBox.NomalColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.FilterComboBox.onHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.FilterComboBox.selectedIndex = -1;
            this.FilterComboBox.Size = new System.Drawing.Size(217, 35);
            this.FilterComboBox.TabIndex = 28;
            this.FilterComboBox.onItemSelected += new System.EventHandler(this.FilterComboBox_onItemSelected);
            // 
            // AccountantWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.Controls.Add(this.FilterComboBox);
            this.Controls.Add(this.BrowseBTN);
            this.Controls.Add(this.AccountmentSum);
            this.Controls.Add(this.ClearBTN);
            this.Controls.Add(this.LogoutBTN);
            this.Controls.Add(this.Accountantgridview);
            this.Name = "AccountantWindow";
            this.Size = new System.Drawing.Size(606, 533);
            this.Load += new System.EventHandler(this.AccountantWindow_Load);
            ((System.ComponentModel.ISupportInitialize)(this.Accountantgridview)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Bunifu.Framework.UI.BunifuCustomDataGrid Accountantgridview;
        private Bunifu.Framework.UI.BunifuFlatButton LogoutBTN;
        private Bunifu.Framework.UI.BunifuFlatButton ClearBTN;
        private Bunifu.Framework.UI.BunifuFlatButton AccountmentSum;
        private Bunifu.Framework.UI.BunifuFlatButton BrowseBTN;
        private Bunifu.Framework.UI.BunifuDropdown FilterComboBox;
    }
}
