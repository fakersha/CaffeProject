namespace CafeProject
{
    partial class ManagerNewProducts
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ManagerNewProducts));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.BackBTN = new Bunifu.Framework.UI.BunifuFlatButton();
            this.NameBox = new Bunifu.Framework.UI.BunifuMaterialTextbox();
            this.PriceBox = new Bunifu.Framework.UI.BunifuMaterialTextbox();
            this.IDBox = new Bunifu.Framework.UI.BunifuMaterialTextbox();
            this.TypeBox = new Bunifu.Framework.UI.BunifuDropdown();
            this.AddBTN = new Bunifu.Framework.UI.BunifuFlatButton();
            this.MenuGridView = new Bunifu.Framework.UI.BunifuCustomDataGrid();
            this.bunifuFlatButton1 = new Bunifu.Framework.UI.BunifuFlatButton();
            this.bunifuFlatButton2 = new Bunifu.Framework.UI.BunifuFlatButton();
            this.RemoveClick = new Bunifu.Framework.UI.BunifuFlatButton();
            this.SaveBTN = new Bunifu.Framework.UI.BunifuFlatButton();
            ((System.ComponentModel.ISupportInitialize)(this.MenuGridView)).BeginInit();
            this.SuspendLayout();
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
            this.BackBTN.TabIndex = 25;
            this.BackBTN.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.BackBTN.Textcolor = System.Drawing.Color.White;
            this.BackBTN.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BackBTN.Click += new System.EventHandler(this.BackBTN_Click);
            // 
            // NameBox
            // 
            this.NameBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.NameBox.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.NameBox.Font = new System.Drawing.Font("Century Gothic", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.NameBox.ForeColor = System.Drawing.Color.White;
            this.NameBox.HintForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(176)))), ((int)(((byte)(183)))), ((int)(((byte)(177)))));
            this.NameBox.HintText = "დასახელება";
            this.NameBox.isPassword = false;
            this.NameBox.LineFocusedColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(100)))), ((int)(((byte)(15)))));
            this.NameBox.LineIdleColor = System.Drawing.Color.Gray;
            this.NameBox.LineMouseHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(100)))), ((int)(((byte)(15)))));
            this.NameBox.LineThickness = 4;
            this.NameBox.Location = new System.Drawing.Point(90, 80);
            this.NameBox.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.NameBox.Name = "NameBox";
            this.NameBox.Size = new System.Drawing.Size(268, 47);
            this.NameBox.TabIndex = 26;
            this.NameBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Left;
            this.NameBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.NameBox_KeyDown);
            // 
            // PriceBox
            // 
            this.PriceBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.PriceBox.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.PriceBox.Font = new System.Drawing.Font("Century Gothic", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.PriceBox.ForeColor = System.Drawing.Color.White;
            this.PriceBox.HintForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(176)))), ((int)(((byte)(183)))), ((int)(((byte)(177)))));
            this.PriceBox.HintText = "ფასი";
            this.PriceBox.isPassword = false;
            this.PriceBox.LineFocusedColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(100)))), ((int)(((byte)(15)))));
            this.PriceBox.LineIdleColor = System.Drawing.Color.Gray;
            this.PriceBox.LineMouseHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(100)))), ((int)(((byte)(15)))));
            this.PriceBox.LineThickness = 4;
            this.PriceBox.Location = new System.Drawing.Point(90, 137);
            this.PriceBox.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.PriceBox.Name = "PriceBox";
            this.PriceBox.Size = new System.Drawing.Size(268, 47);
            this.PriceBox.TabIndex = 27;
            this.PriceBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Left;
            this.PriceBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.PriceBox_KeyDown);
            // 
            // IDBox
            // 
            this.IDBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.IDBox.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.IDBox.Font = new System.Drawing.Font("Century Gothic", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.IDBox.ForeColor = System.Drawing.Color.White;
            this.IDBox.HintForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(176)))), ((int)(((byte)(183)))), ((int)(((byte)(177)))));
            this.IDBox.HintText = "ID";
            this.IDBox.isPassword = false;
            this.IDBox.LineFocusedColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(100)))), ((int)(((byte)(15)))));
            this.IDBox.LineIdleColor = System.Drawing.Color.Gray;
            this.IDBox.LineMouseHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(100)))), ((int)(((byte)(15)))));
            this.IDBox.LineThickness = 4;
            this.IDBox.Location = new System.Drawing.Point(90, 195);
            this.IDBox.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.IDBox.Name = "IDBox";
            this.IDBox.Size = new System.Drawing.Size(268, 47);
            this.IDBox.TabIndex = 28;
            this.IDBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Left;
            this.IDBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.IDBox_KeyDown);
            // 
            // TypeBox
            // 
            this.TypeBox.BackColor = System.Drawing.Color.Transparent;
            this.TypeBox.BorderRadius = 3;
            this.TypeBox.ForeColor = System.Drawing.Color.White;
            this.TypeBox.Items = new string[0];
            this.TypeBox.Location = new System.Drawing.Point(90, 253);
            this.TypeBox.Name = "TypeBox";
            this.TypeBox.NomalColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.TypeBox.onHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.TypeBox.selectedIndex = -1;
            this.TypeBox.Size = new System.Drawing.Size(217, 40);
            this.TypeBox.TabIndex = 29;
            // 
            // AddBTN
            // 
            this.AddBTN.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.AddBTN.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.AddBTN.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.AddBTN.BorderRadius = 0;
            this.AddBTN.ButtonText = "დამატება";
            this.AddBTN.Cursor = System.Windows.Forms.Cursors.Hand;
            this.AddBTN.DisabledColor = System.Drawing.Color.Gray;
            this.AddBTN.Iconcolor = System.Drawing.Color.Transparent;
            this.AddBTN.Iconimage = ((System.Drawing.Image)(resources.GetObject("AddBTN.Iconimage")));
            this.AddBTN.Iconimage_right = null;
            this.AddBTN.Iconimage_right_Selected = null;
            this.AddBTN.Iconimage_Selected = null;
            this.AddBTN.IconMarginLeft = 0;
            this.AddBTN.IconMarginRight = 0;
            this.AddBTN.IconRightVisible = true;
            this.AddBTN.IconRightZoom = 0D;
            this.AddBTN.IconVisible = true;
            this.AddBTN.IconZoom = 50D;
            this.AddBTN.IsTab = false;
            this.AddBTN.Location = new System.Drawing.Point(122, 322);
            this.AddBTN.Name = "AddBTN";
            this.AddBTN.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.AddBTN.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.AddBTN.OnHoverTextColor = System.Drawing.Color.White;
            this.AddBTN.selected = false;
            this.AddBTN.Size = new System.Drawing.Size(204, 52);
            this.AddBTN.TabIndex = 30;
            this.AddBTN.Text = "დამატება";
            this.AddBTN.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.AddBTN.Textcolor = System.Drawing.Color.White;
            this.AddBTN.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.AddBTN.Click += new System.EventHandler(this.AddBTN_Click);
            // 
            // MenuGridView
            // 
            this.MenuGridView.AllowUserToDeleteRows = false;
            this.MenuGridView.AllowUserToOrderColumns = true;
            this.MenuGridView.AllowUserToResizeColumns = false;
            this.MenuGridView.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.MenuGridView.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.MenuGridView.BackgroundColor = System.Drawing.Color.Gainsboro;
            this.MenuGridView.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.MenuGridView.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Century Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.MenuGridView.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.MenuGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.MenuGridView.DoubleBuffered = true;
            this.MenuGridView.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.MenuGridView.EnableHeadersVisualStyles = false;
            this.MenuGridView.HeaderBgColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.MenuGridView.HeaderForeColor = System.Drawing.Color.White;
            this.MenuGridView.Location = new System.Drawing.Point(4, 450);
            this.MenuGridView.Name = "MenuGridView";
            this.MenuGridView.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            this.MenuGridView.Size = new System.Drawing.Size(441, 150);
            this.MenuGridView.TabIndex = 31;
            // 
            // bunifuFlatButton1
            // 
            this.bunifuFlatButton1.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.bunifuFlatButton1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))), ((int)(((byte)(80)))));
            this.bunifuFlatButton1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.bunifuFlatButton1.BorderRadius = 7;
            this.bunifuFlatButton1.ButtonText = "";
            this.bunifuFlatButton1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.bunifuFlatButton1.DisabledColor = System.Drawing.Color.Gray;
            this.bunifuFlatButton1.Iconcolor = System.Drawing.Color.Transparent;
            this.bunifuFlatButton1.Iconimage = ((System.Drawing.Image)(resources.GetObject("bunifuFlatButton1.Iconimage")));
            this.bunifuFlatButton1.Iconimage_right = null;
            this.bunifuFlatButton1.Iconimage_right_Selected = null;
            this.bunifuFlatButton1.Iconimage_Selected = null;
            this.bunifuFlatButton1.IconMarginLeft = 0;
            this.bunifuFlatButton1.IconMarginRight = 0;
            this.bunifuFlatButton1.IconRightVisible = true;
            this.bunifuFlatButton1.IconRightZoom = 0D;
            this.bunifuFlatButton1.IconVisible = true;
            this.bunifuFlatButton1.IconZoom = 50D;
            this.bunifuFlatButton1.IsTab = false;
            this.bunifuFlatButton1.Location = new System.Drawing.Point(205, 396);
            this.bunifuFlatButton1.Name = "bunifuFlatButton1";
            this.bunifuFlatButton1.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))), ((int)(((byte)(80)))));
            this.bunifuFlatButton1.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.bunifuFlatButton1.OnHoverTextColor = System.Drawing.Color.White;
            this.bunifuFlatButton1.selected = false;
            this.bunifuFlatButton1.Size = new System.Drawing.Size(36, 39);
            this.bunifuFlatButton1.TabIndex = 32;
            this.bunifuFlatButton1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.bunifuFlatButton1.Textcolor = System.Drawing.Color.White;
            this.bunifuFlatButton1.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bunifuFlatButton1.Click += new System.EventHandler(this.bunifuFlatButton1_Click);
            // 
            // bunifuFlatButton2
            // 
            this.bunifuFlatButton2.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.bunifuFlatButton2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))), ((int)(((byte)(80)))));
            this.bunifuFlatButton2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.bunifuFlatButton2.BorderRadius = 7;
            this.bunifuFlatButton2.ButtonText = "";
            this.bunifuFlatButton2.Cursor = System.Windows.Forms.Cursors.Hand;
            this.bunifuFlatButton2.DisabledColor = System.Drawing.Color.Gray;
            this.bunifuFlatButton2.Iconcolor = System.Drawing.Color.Transparent;
            this.bunifuFlatButton2.Iconimage = ((System.Drawing.Image)(resources.GetObject("bunifuFlatButton2.Iconimage")));
            this.bunifuFlatButton2.Iconimage_right = null;
            this.bunifuFlatButton2.Iconimage_right_Selected = null;
            this.bunifuFlatButton2.Iconimage_Selected = null;
            this.bunifuFlatButton2.IconMarginLeft = 0;
            this.bunifuFlatButton2.IconMarginRight = 0;
            this.bunifuFlatButton2.IconRightVisible = true;
            this.bunifuFlatButton2.IconRightZoom = 0D;
            this.bunifuFlatButton2.IconVisible = true;
            this.bunifuFlatButton2.IconZoom = 50D;
            this.bunifuFlatButton2.IsTab = false;
            this.bunifuFlatButton2.Location = new System.Drawing.Point(205, 396);
            this.bunifuFlatButton2.Name = "bunifuFlatButton2";
            this.bunifuFlatButton2.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))), ((int)(((byte)(80)))));
            this.bunifuFlatButton2.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.bunifuFlatButton2.OnHoverTextColor = System.Drawing.Color.White;
            this.bunifuFlatButton2.selected = false;
            this.bunifuFlatButton2.Size = new System.Drawing.Size(36, 39);
            this.bunifuFlatButton2.TabIndex = 33;
            this.bunifuFlatButton2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.bunifuFlatButton2.Textcolor = System.Drawing.Color.White;
            this.bunifuFlatButton2.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bunifuFlatButton2.Click += new System.EventHandler(this.bunifuFlatButton2_Click);
            // 
            // RemoveClick
            // 
            this.RemoveClick.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.RemoveClick.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.RemoveClick.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.RemoveClick.BorderRadius = 0;
            this.RemoveClick.ButtonText = "წაშლა";
            this.RemoveClick.Cursor = System.Windows.Forms.Cursors.Hand;
            this.RemoveClick.DisabledColor = System.Drawing.Color.Gray;
            this.RemoveClick.Iconcolor = System.Drawing.Color.Transparent;
            this.RemoveClick.Iconimage = ((System.Drawing.Image)(resources.GetObject("RemoveClick.Iconimage")));
            this.RemoveClick.Iconimage_right = null;
            this.RemoveClick.Iconimage_right_Selected = null;
            this.RemoveClick.Iconimage_Selected = null;
            this.RemoveClick.IconMarginLeft = 0;
            this.RemoveClick.IconMarginRight = 0;
            this.RemoveClick.IconRightVisible = true;
            this.RemoveClick.IconRightZoom = 0D;
            this.RemoveClick.IconVisible = true;
            this.RemoveClick.IconZoom = 50D;
            this.RemoveClick.IsTab = false;
            this.RemoveClick.Location = new System.Drawing.Point(0, 607);
            this.RemoveClick.Name = "RemoveClick";
            this.RemoveClick.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.RemoveClick.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.RemoveClick.OnHoverTextColor = System.Drawing.Color.White;
            this.RemoveClick.selected = false;
            this.RemoveClick.Size = new System.Drawing.Size(224, 44);
            this.RemoveClick.TabIndex = 34;
            this.RemoveClick.Text = "წაშლა";
            this.RemoveClick.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.RemoveClick.Textcolor = System.Drawing.Color.White;
            this.RemoveClick.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RemoveClick.Click += new System.EventHandler(this.RemoveClick_Click);
            // 
            // SaveBTN
            // 
            this.SaveBTN.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.SaveBTN.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.SaveBTN.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.SaveBTN.BorderRadius = 0;
            this.SaveBTN.ButtonText = "შენახვა";
            this.SaveBTN.Cursor = System.Windows.Forms.Cursors.Hand;
            this.SaveBTN.DisabledColor = System.Drawing.Color.Gray;
            this.SaveBTN.Iconcolor = System.Drawing.Color.Transparent;
            this.SaveBTN.Iconimage = ((System.Drawing.Image)(resources.GetObject("SaveBTN.Iconimage")));
            this.SaveBTN.Iconimage_right = null;
            this.SaveBTN.Iconimage_right_Selected = null;
            this.SaveBTN.Iconimage_Selected = null;
            this.SaveBTN.IconMarginLeft = 0;
            this.SaveBTN.IconMarginRight = 0;
            this.SaveBTN.IconRightVisible = true;
            this.SaveBTN.IconRightZoom = 0D;
            this.SaveBTN.IconVisible = true;
            this.SaveBTN.IconZoom = 50D;
            this.SaveBTN.IsTab = false;
            this.SaveBTN.Location = new System.Drawing.Point(225, 607);
            this.SaveBTN.Name = "SaveBTN";
            this.SaveBTN.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.SaveBTN.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.SaveBTN.OnHoverTextColor = System.Drawing.Color.White;
            this.SaveBTN.selected = false;
            this.SaveBTN.Size = new System.Drawing.Size(224, 44);
            this.SaveBTN.TabIndex = 40;
            this.SaveBTN.Text = "შენახვა";
            this.SaveBTN.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.SaveBTN.Textcolor = System.Drawing.Color.White;
            this.SaveBTN.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SaveBTN.Click += new System.EventHandler(this.SaveBTN_Click);
            // 
            // ManagerNewProducts
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.Controls.Add(this.SaveBTN);
            this.Controls.Add(this.RemoveClick);
            this.Controls.Add(this.bunifuFlatButton2);
            this.Controls.Add(this.bunifuFlatButton1);
            this.Controls.Add(this.MenuGridView);
            this.Controls.Add(this.AddBTN);
            this.Controls.Add(this.TypeBox);
            this.Controls.Add(this.IDBox);
            this.Controls.Add(this.PriceBox);
            this.Controls.Add(this.NameBox);
            this.Controls.Add(this.BackBTN);
            this.Name = "ManagerNewProducts";
            this.Size = new System.Drawing.Size(448, 650);
            this.Load += new System.EventHandler(this.ManagerNewProducts_Load);
            ((System.ComponentModel.ISupportInitialize)(this.MenuGridView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Bunifu.Framework.UI.BunifuFlatButton BackBTN;
        private Bunifu.Framework.UI.BunifuMaterialTextbox NameBox;
        private Bunifu.Framework.UI.BunifuMaterialTextbox PriceBox;
        private Bunifu.Framework.UI.BunifuMaterialTextbox IDBox;
        private Bunifu.Framework.UI.BunifuDropdown TypeBox;
        private Bunifu.Framework.UI.BunifuFlatButton AddBTN;
        private Bunifu.Framework.UI.BunifuCustomDataGrid MenuGridView;
        private Bunifu.Framework.UI.BunifuFlatButton bunifuFlatButton1;
        private Bunifu.Framework.UI.BunifuFlatButton bunifuFlatButton2;
        private Bunifu.Framework.UI.BunifuFlatButton RemoveClick;
        private Bunifu.Framework.UI.BunifuFlatButton SaveBTN;
    }
}
