namespace CafeProject
{
    partial class WaiterWindow
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WaiterWindow));
            this.dataGridView1 = new Bunifu.Framework.UI.BunifuCustomDataGrid();
            this.dataGridView3 = new Bunifu.Framework.UI.BunifuCustomDataGrid();
            this.Delete = new Bunifu.Framework.UI.BunifuFlatButton();
            this.RemoveClick = new Bunifu.Framework.UI.BunifuFlatButton();
            this.sumOfOrder = new Bunifu.Framework.UI.BunifuFlatButton();
            this.Table_1 = new Bunifu.Framework.UI.BunifuFlatButton();
            this.Table_2 = new Bunifu.Framework.UI.BunifuFlatButton();
            this.Table_3 = new Bunifu.Framework.UI.BunifuFlatButton();
            this.Table_4 = new Bunifu.Framework.UI.BunifuFlatButton();
            this.Table_5 = new Bunifu.Framework.UI.BunifuFlatButton();
            this.Table_6 = new Bunifu.Framework.UI.BunifuFlatButton();
            this.Table_7 = new Bunifu.Framework.UI.BunifuFlatButton();
            this.Table_8 = new Bunifu.Framework.UI.BunifuFlatButton();
            this.TypeBox = new Bunifu.Framework.UI.BunifuDropdown();
            this.ResultCombobox = new Bunifu.Framework.UI.BunifuDropdown();
            this.PayTypeComboBox = new Bunifu.Framework.UI.BunifuDropdown();
            this.TableClearBTN = new Bunifu.Framework.UI.BunifuFlatButton();
            this.Save = new Bunifu.Framework.UI.BunifuFlatButton();
            this.ChangeBox = new Bunifu.Framework.UI.BunifuMaterialTextbox();
            this.DescriptionBox = new Bunifu.Framework.UI.BunifuMaterialTextbox();
            this.BackBTN = new Bunifu.Framework.UI.BunifuFlatButton();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AllowUserToOrderColumns = true;
            this.dataGridView1.AllowUserToResizeColumns = false;
            this.dataGridView1.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.dataGridView1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView1.BackgroundColor = System.Drawing.Color.Gainsboro;
            this.dataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dataGridView1.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Century Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.DoubleBuffered = true;
            this.dataGridView1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.dataGridView1.EnableHeadersVisualStyles = false;
            this.dataGridView1.HeaderBgColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.dataGridView1.HeaderForeColor = System.Drawing.Color.White;
            this.dataGridView1.Location = new System.Drawing.Point(3, 3);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            this.dataGridView1.Size = new System.Drawing.Size(441, 150);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.RowHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridView1_RowHeaderMouseClick);
            // 
            // dataGridView3
            // 
            this.dataGridView3.AllowUserToDeleteRows = false;
            this.dataGridView3.AllowUserToResizeColumns = false;
            this.dataGridView3.AllowUserToResizeRows = false;
            dataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.dataGridView3.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle3;
            this.dataGridView3.BackgroundColor = System.Drawing.Color.Gainsboro;
            this.dataGridView3.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dataGridView3.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Century Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView3.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle4;
            this.dataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView3.DoubleBuffered = true;
            this.dataGridView3.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.dataGridView3.EnableHeadersVisualStyles = false;
            this.dataGridView3.HeaderBgColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.dataGridView3.HeaderForeColor = System.Drawing.Color.White;
            this.dataGridView3.Location = new System.Drawing.Point(648, 3);
            this.dataGridView3.Name = "dataGridView3";
            this.dataGridView3.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            this.dataGridView3.Size = new System.Drawing.Size(541, 150);
            this.dataGridView3.TabIndex = 1;
            // 
            // Delete
            // 
            this.Delete.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.Delete.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.Delete.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.Delete.BorderRadius = 0;
            this.Delete.ButtonText = "ამოკლება";
            this.Delete.Cursor = System.Windows.Forms.Cursors.Hand;
            this.Delete.DisabledColor = System.Drawing.Color.Gray;
            this.Delete.Iconcolor = System.Drawing.Color.Transparent;
            this.Delete.Iconimage = ((System.Drawing.Image)(resources.GetObject("Delete.Iconimage")));
            this.Delete.Iconimage_right = null;
            this.Delete.Iconimage_right_Selected = null;
            this.Delete.Iconimage_Selected = null;
            this.Delete.IconMarginLeft = 0;
            this.Delete.IconMarginRight = 0;
            this.Delete.IconRightVisible = true;
            this.Delete.IconRightZoom = 0D;
            this.Delete.IconVisible = true;
            this.Delete.IconZoom = 50D;
            this.Delete.IsTab = false;
            this.Delete.Location = new System.Drawing.Point(464, 19);
            this.Delete.Name = "Delete";
            this.Delete.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.Delete.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.Delete.OnHoverTextColor = System.Drawing.Color.White;
            this.Delete.selected = false;
            this.Delete.Size = new System.Drawing.Size(163, 36);
            this.Delete.TabIndex = 2;
            this.Delete.Text = "ამოკლება";
            this.Delete.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Delete.Textcolor = System.Drawing.Color.White;
            this.Delete.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Delete.Click += new System.EventHandler(this.Delete_Click);
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
            this.RemoveClick.Location = new System.Drawing.Point(464, 61);
            this.RemoveClick.Name = "RemoveClick";
            this.RemoveClick.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.RemoveClick.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.RemoveClick.OnHoverTextColor = System.Drawing.Color.White;
            this.RemoveClick.selected = false;
            this.RemoveClick.Size = new System.Drawing.Size(163, 36);
            this.RemoveClick.TabIndex = 3;
            this.RemoveClick.Text = "წაშლა";
            this.RemoveClick.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.RemoveClick.Textcolor = System.Drawing.Color.White;
            this.RemoveClick.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RemoveClick.Click += new System.EventHandler(this.RemoveClick_Click);
            // 
            // sumOfOrder
            // 
            this.sumOfOrder.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.sumOfOrder.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.sumOfOrder.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.sumOfOrder.BorderRadius = 0;
            this.sumOfOrder.ButtonText = "ანგარიში";
            this.sumOfOrder.Cursor = System.Windows.Forms.Cursors.Hand;
            this.sumOfOrder.DisabledColor = System.Drawing.Color.Gray;
            this.sumOfOrder.Iconcolor = System.Drawing.Color.Transparent;
            this.sumOfOrder.Iconimage = ((System.Drawing.Image)(resources.GetObject("sumOfOrder.Iconimage")));
            this.sumOfOrder.Iconimage_right = null;
            this.sumOfOrder.Iconimage_right_Selected = null;
            this.sumOfOrder.Iconimage_Selected = null;
            this.sumOfOrder.IconMarginLeft = 0;
            this.sumOfOrder.IconMarginRight = 0;
            this.sumOfOrder.IconRightVisible = true;
            this.sumOfOrder.IconRightZoom = 0D;
            this.sumOfOrder.IconVisible = true;
            this.sumOfOrder.IconZoom = 50D;
            this.sumOfOrder.IsTab = false;
            this.sumOfOrder.Location = new System.Drawing.Point(464, 103);
            this.sumOfOrder.Name = "sumOfOrder";
            this.sumOfOrder.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.sumOfOrder.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.sumOfOrder.OnHoverTextColor = System.Drawing.Color.White;
            this.sumOfOrder.selected = false;
            this.sumOfOrder.Size = new System.Drawing.Size(163, 36);
            this.sumOfOrder.TabIndex = 4;
            this.sumOfOrder.Text = "ანგარიში";
            this.sumOfOrder.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.sumOfOrder.Textcolor = System.Drawing.Color.White;
            this.sumOfOrder.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.sumOfOrder.Click += new System.EventHandler(this.sumOfOrder_Click);
            // 
            // Table_1
            // 
            this.Table_1.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.Table_1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.Table_1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.Table_1.BorderRadius = 0;
            this.Table_1.ButtonText = "";
            this.Table_1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.Table_1.DisabledColor = System.Drawing.Color.Gray;
            this.Table_1.Iconcolor = System.Drawing.Color.Transparent;
            this.Table_1.Iconimage = ((System.Drawing.Image)(resources.GetObject("Table_1.Iconimage")));
            this.Table_1.Iconimage_right = null;
            this.Table_1.Iconimage_right_Selected = null;
            this.Table_1.Iconimage_Selected = null;
            this.Table_1.IconMarginLeft = 0;
            this.Table_1.IconMarginRight = 0;
            this.Table_1.IconRightVisible = true;
            this.Table_1.IconRightZoom = 30D;
            this.Table_1.IconVisible = true;
            this.Table_1.IconZoom = 40D;
            this.Table_1.IsTab = false;
            this.Table_1.Location = new System.Drawing.Point(3, 319);
            this.Table_1.Name = "Table_1";
            this.Table_1.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.Table_1.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.Table_1.OnHoverTextColor = System.Drawing.Color.White;
            this.Table_1.selected = false;
            this.Table_1.Size = new System.Drawing.Size(84, 83);
            this.Table_1.TabIndex = 5;
            this.Table_1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Table_1.Textcolor = System.Drawing.Color.White;
            this.Table_1.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Table_1.Click += new System.EventHandler(this.Table_1_Click);
            // 
            // Table_2
            // 
            this.Table_2.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.Table_2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.Table_2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.Table_2.BorderRadius = 0;
            this.Table_2.ButtonText = "";
            this.Table_2.Cursor = System.Windows.Forms.Cursors.Hand;
            this.Table_2.DisabledColor = System.Drawing.Color.Gray;
            this.Table_2.Iconcolor = System.Drawing.Color.Transparent;
            this.Table_2.Iconimage = ((System.Drawing.Image)(resources.GetObject("Table_2.Iconimage")));
            this.Table_2.Iconimage_right = null;
            this.Table_2.Iconimage_right_Selected = null;
            this.Table_2.Iconimage_Selected = null;
            this.Table_2.IconMarginLeft = 0;
            this.Table_2.IconMarginRight = 0;
            this.Table_2.IconRightVisible = true;
            this.Table_2.IconRightZoom = 30D;
            this.Table_2.IconVisible = true;
            this.Table_2.IconZoom = 40D;
            this.Table_2.IsTab = false;
            this.Table_2.Location = new System.Drawing.Point(93, 319);
            this.Table_2.Name = "Table_2";
            this.Table_2.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.Table_2.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.Table_2.OnHoverTextColor = System.Drawing.Color.White;
            this.Table_2.selected = false;
            this.Table_2.Size = new System.Drawing.Size(84, 83);
            this.Table_2.TabIndex = 6;
            this.Table_2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Table_2.Textcolor = System.Drawing.Color.White;
            this.Table_2.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Table_2.Click += new System.EventHandler(this.Table_2_Click);
            // 
            // Table_3
            // 
            this.Table_3.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.Table_3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.Table_3.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.Table_3.BorderRadius = 0;
            this.Table_3.ButtonText = "";
            this.Table_3.Cursor = System.Windows.Forms.Cursors.Hand;
            this.Table_3.DisabledColor = System.Drawing.Color.Gray;
            this.Table_3.Iconcolor = System.Drawing.Color.Transparent;
            this.Table_3.Iconimage = ((System.Drawing.Image)(resources.GetObject("Table_3.Iconimage")));
            this.Table_3.Iconimage_right = null;
            this.Table_3.Iconimage_right_Selected = null;
            this.Table_3.Iconimage_Selected = null;
            this.Table_3.IconMarginLeft = 0;
            this.Table_3.IconMarginRight = 0;
            this.Table_3.IconRightVisible = true;
            this.Table_3.IconRightZoom = 30D;
            this.Table_3.IconVisible = true;
            this.Table_3.IconZoom = 40D;
            this.Table_3.IsTab = false;
            this.Table_3.Location = new System.Drawing.Point(183, 319);
            this.Table_3.Name = "Table_3";
            this.Table_3.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.Table_3.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.Table_3.OnHoverTextColor = System.Drawing.Color.White;
            this.Table_3.selected = false;
            this.Table_3.Size = new System.Drawing.Size(84, 83);
            this.Table_3.TabIndex = 7;
            this.Table_3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Table_3.Textcolor = System.Drawing.Color.White;
            this.Table_3.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Table_3.Click += new System.EventHandler(this.Table_3_Click);
            // 
            // Table_4
            // 
            this.Table_4.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.Table_4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.Table_4.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.Table_4.BorderRadius = 0;
            this.Table_4.ButtonText = "";
            this.Table_4.Cursor = System.Windows.Forms.Cursors.Hand;
            this.Table_4.DisabledColor = System.Drawing.Color.Gray;
            this.Table_4.Iconcolor = System.Drawing.Color.Transparent;
            this.Table_4.Iconimage = ((System.Drawing.Image)(resources.GetObject("Table_4.Iconimage")));
            this.Table_4.Iconimage_right = null;
            this.Table_4.Iconimage_right_Selected = null;
            this.Table_4.Iconimage_Selected = null;
            this.Table_4.IconMarginLeft = 0;
            this.Table_4.IconMarginRight = 0;
            this.Table_4.IconRightVisible = true;
            this.Table_4.IconRightZoom = 30D;
            this.Table_4.IconVisible = true;
            this.Table_4.IconZoom = 40D;
            this.Table_4.IsTab = false;
            this.Table_4.Location = new System.Drawing.Point(273, 319);
            this.Table_4.Name = "Table_4";
            this.Table_4.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.Table_4.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.Table_4.OnHoverTextColor = System.Drawing.Color.White;
            this.Table_4.selected = false;
            this.Table_4.Size = new System.Drawing.Size(84, 83);
            this.Table_4.TabIndex = 8;
            this.Table_4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Table_4.Textcolor = System.Drawing.Color.White;
            this.Table_4.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Table_4.Click += new System.EventHandler(this.Table_4_Click);
            // 
            // Table_5
            // 
            this.Table_5.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.Table_5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.Table_5.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.Table_5.BorderRadius = 0;
            this.Table_5.ButtonText = "";
            this.Table_5.Cursor = System.Windows.Forms.Cursors.Hand;
            this.Table_5.DisabledColor = System.Drawing.Color.Gray;
            this.Table_5.Iconcolor = System.Drawing.Color.Transparent;
            this.Table_5.Iconimage = ((System.Drawing.Image)(resources.GetObject("Table_5.Iconimage")));
            this.Table_5.Iconimage_right = null;
            this.Table_5.Iconimage_right_Selected = null;
            this.Table_5.Iconimage_Selected = null;
            this.Table_5.IconMarginLeft = 0;
            this.Table_5.IconMarginRight = 0;
            this.Table_5.IconRightVisible = true;
            this.Table_5.IconRightZoom = 30D;
            this.Table_5.IconVisible = true;
            this.Table_5.IconZoom = 40D;
            this.Table_5.IsTab = false;
            this.Table_5.Location = new System.Drawing.Point(3, 408);
            this.Table_5.Name = "Table_5";
            this.Table_5.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.Table_5.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.Table_5.OnHoverTextColor = System.Drawing.Color.White;
            this.Table_5.selected = false;
            this.Table_5.Size = new System.Drawing.Size(84, 83);
            this.Table_5.TabIndex = 9;
            this.Table_5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Table_5.Textcolor = System.Drawing.Color.White;
            this.Table_5.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Table_5.Click += new System.EventHandler(this.Table_5_Click);
            // 
            // Table_6
            // 
            this.Table_6.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.Table_6.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.Table_6.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.Table_6.BorderRadius = 0;
            this.Table_6.ButtonText = "";
            this.Table_6.Cursor = System.Windows.Forms.Cursors.Hand;
            this.Table_6.DisabledColor = System.Drawing.Color.Gray;
            this.Table_6.Iconcolor = System.Drawing.Color.Transparent;
            this.Table_6.Iconimage = ((System.Drawing.Image)(resources.GetObject("Table_6.Iconimage")));
            this.Table_6.Iconimage_right = null;
            this.Table_6.Iconimage_right_Selected = null;
            this.Table_6.Iconimage_Selected = null;
            this.Table_6.IconMarginLeft = 0;
            this.Table_6.IconMarginRight = 0;
            this.Table_6.IconRightVisible = true;
            this.Table_6.IconRightZoom = 30D;
            this.Table_6.IconVisible = true;
            this.Table_6.IconZoom = 40D;
            this.Table_6.IsTab = false;
            this.Table_6.Location = new System.Drawing.Point(93, 408);
            this.Table_6.Name = "Table_6";
            this.Table_6.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.Table_6.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.Table_6.OnHoverTextColor = System.Drawing.Color.White;
            this.Table_6.selected = false;
            this.Table_6.Size = new System.Drawing.Size(84, 83);
            this.Table_6.TabIndex = 10;
            this.Table_6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Table_6.Textcolor = System.Drawing.Color.White;
            this.Table_6.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Table_6.Click += new System.EventHandler(this.Table_6_Click);
            // 
            // Table_7
            // 
            this.Table_7.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.Table_7.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.Table_7.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.Table_7.BorderRadius = 0;
            this.Table_7.ButtonText = "";
            this.Table_7.Cursor = System.Windows.Forms.Cursors.Hand;
            this.Table_7.DisabledColor = System.Drawing.Color.Gray;
            this.Table_7.Iconcolor = System.Drawing.Color.Transparent;
            this.Table_7.Iconimage = ((System.Drawing.Image)(resources.GetObject("Table_7.Iconimage")));
            this.Table_7.Iconimage_right = null;
            this.Table_7.Iconimage_right_Selected = null;
            this.Table_7.Iconimage_Selected = null;
            this.Table_7.IconMarginLeft = 0;
            this.Table_7.IconMarginRight = 0;
            this.Table_7.IconRightVisible = true;
            this.Table_7.IconRightZoom = 30D;
            this.Table_7.IconVisible = true;
            this.Table_7.IconZoom = 40D;
            this.Table_7.IsTab = false;
            this.Table_7.Location = new System.Drawing.Point(183, 408);
            this.Table_7.Name = "Table_7";
            this.Table_7.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.Table_7.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.Table_7.OnHoverTextColor = System.Drawing.Color.White;
            this.Table_7.selected = false;
            this.Table_7.Size = new System.Drawing.Size(84, 83);
            this.Table_7.TabIndex = 11;
            this.Table_7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Table_7.Textcolor = System.Drawing.Color.White;
            this.Table_7.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Table_7.Click += new System.EventHandler(this.Table_7_Click);
            // 
            // Table_8
            // 
            this.Table_8.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.Table_8.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.Table_8.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.Table_8.BorderRadius = 0;
            this.Table_8.ButtonText = "";
            this.Table_8.Cursor = System.Windows.Forms.Cursors.Hand;
            this.Table_8.DisabledColor = System.Drawing.Color.Gray;
            this.Table_8.Iconcolor = System.Drawing.Color.Transparent;
            this.Table_8.Iconimage = ((System.Drawing.Image)(resources.GetObject("Table_8.Iconimage")));
            this.Table_8.Iconimage_right = null;
            this.Table_8.Iconimage_right_Selected = null;
            this.Table_8.Iconimage_Selected = null;
            this.Table_8.IconMarginLeft = 0;
            this.Table_8.IconMarginRight = 0;
            this.Table_8.IconRightVisible = true;
            this.Table_8.IconRightZoom = 30D;
            this.Table_8.IconVisible = true;
            this.Table_8.IconZoom = 40D;
            this.Table_8.IsTab = false;
            this.Table_8.Location = new System.Drawing.Point(273, 408);
            this.Table_8.Name = "Table_8";
            this.Table_8.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.Table_8.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.Table_8.OnHoverTextColor = System.Drawing.Color.White;
            this.Table_8.selected = false;
            this.Table_8.Size = new System.Drawing.Size(84, 83);
            this.Table_8.TabIndex = 12;
            this.Table_8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Table_8.Textcolor = System.Drawing.Color.White;
            this.Table_8.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Table_8.Click += new System.EventHandler(this.Table_8_Click);
            // 
            // TypeBox
            // 
            this.TypeBox.BackColor = System.Drawing.Color.Transparent;
            this.TypeBox.BorderRadius = 3;
            this.TypeBox.ForeColor = System.Drawing.Color.White;
            this.TypeBox.Items = new string[0];
            this.TypeBox.Location = new System.Drawing.Point(3, 159);
            this.TypeBox.Name = "TypeBox";
            this.TypeBox.NomalColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.TypeBox.onHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.TypeBox.selectedIndex = -1;
            this.TypeBox.Size = new System.Drawing.Size(217, 35);
            this.TypeBox.TabIndex = 13;
            this.TypeBox.onItemSelected += new System.EventHandler(this.TypeBox_onItemSelected);
            // 
            // ResultCombobox
            // 
            this.ResultCombobox.BackColor = System.Drawing.Color.Transparent;
            this.ResultCombobox.BorderRadius = 3;
            this.ResultCombobox.ForeColor = System.Drawing.Color.White;
            this.ResultCombobox.Items = new string[0];
            this.ResultCombobox.Location = new System.Drawing.Point(227, 159);
            this.ResultCombobox.Name = "ResultCombobox";
            this.ResultCombobox.NomalColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.ResultCombobox.onHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.ResultCombobox.selectedIndex = -1;
            this.ResultCombobox.Size = new System.Drawing.Size(217, 35);
            this.ResultCombobox.TabIndex = 14;
            this.ResultCombobox.onItemSelected += new System.EventHandler(this.ResultCombobox_onItemSelected);
            // 
            // PayTypeComboBox
            // 
            this.PayTypeComboBox.BackColor = System.Drawing.Color.Transparent;
            this.PayTypeComboBox.BorderRadius = 3;
            this.PayTypeComboBox.ForeColor = System.Drawing.Color.White;
            this.PayTypeComboBox.Items = new string[0];
            this.PayTypeComboBox.Location = new System.Drawing.Point(648, 159);
            this.PayTypeComboBox.Name = "PayTypeComboBox";
            this.PayTypeComboBox.NomalColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.PayTypeComboBox.onHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.PayTypeComboBox.selectedIndex = -1;
            this.PayTypeComboBox.Size = new System.Drawing.Size(217, 35);
            this.PayTypeComboBox.TabIndex = 15;
            this.PayTypeComboBox.onItemSelected += new System.EventHandler(this.PayTypeComboBox_onItemSelected);
            // 
            // TableClearBTN
            // 
            this.TableClearBTN.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.TableClearBTN.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.TableClearBTN.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.TableClearBTN.BorderRadius = 0;
            this.TableClearBTN.ButtonText = "გასუფთავება";
            this.TableClearBTN.Cursor = System.Windows.Forms.Cursors.Hand;
            this.TableClearBTN.DisabledColor = System.Drawing.Color.Gray;
            this.TableClearBTN.Iconcolor = System.Drawing.Color.Transparent;
            this.TableClearBTN.Iconimage = ((System.Drawing.Image)(resources.GetObject("TableClearBTN.Iconimage")));
            this.TableClearBTN.Iconimage_right = null;
            this.TableClearBTN.Iconimage_right_Selected = null;
            this.TableClearBTN.Iconimage_Selected = null;
            this.TableClearBTN.IconMarginLeft = 0;
            this.TableClearBTN.IconMarginRight = 0;
            this.TableClearBTN.IconRightVisible = true;
            this.TableClearBTN.IconRightZoom = 0D;
            this.TableClearBTN.IconVisible = true;
            this.TableClearBTN.IconZoom = 50D;
            this.TableClearBTN.IsTab = false;
            this.TableClearBTN.Location = new System.Drawing.Point(871, 159);
            this.TableClearBTN.Name = "TableClearBTN";
            this.TableClearBTN.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.TableClearBTN.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.TableClearBTN.OnHoverTextColor = System.Drawing.Color.White;
            this.TableClearBTN.selected = false;
            this.TableClearBTN.Size = new System.Drawing.Size(149, 36);
            this.TableClearBTN.TabIndex = 16;
            this.TableClearBTN.Text = "გასუფთავება";
            this.TableClearBTN.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.TableClearBTN.Textcolor = System.Drawing.Color.White;
            this.TableClearBTN.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TableClearBTN.Click += new System.EventHandler(this.TableClearBTN_Click);
            // 
            // Save
            // 
            this.Save.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.Save.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.Save.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.Save.BorderRadius = 0;
            this.Save.ButtonText = "შენახვა";
            this.Save.Cursor = System.Windows.Forms.Cursors.Hand;
            this.Save.DisabledColor = System.Drawing.Color.Gray;
            this.Save.Iconcolor = System.Drawing.Color.Transparent;
            this.Save.Iconimage = ((System.Drawing.Image)(resources.GetObject("Save.Iconimage")));
            this.Save.Iconimage_right = null;
            this.Save.Iconimage_right_Selected = null;
            this.Save.Iconimage_Selected = null;
            this.Save.IconMarginLeft = 0;
            this.Save.IconMarginRight = 0;
            this.Save.IconRightVisible = true;
            this.Save.IconRightZoom = 0D;
            this.Save.IconVisible = true;
            this.Save.IconZoom = 50D;
            this.Save.IsTab = false;
            this.Save.Location = new System.Drawing.Point(1026, 159);
            this.Save.Name = "Save";
            this.Save.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.Save.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.Save.OnHoverTextColor = System.Drawing.Color.White;
            this.Save.selected = false;
            this.Save.Size = new System.Drawing.Size(149, 36);
            this.Save.TabIndex = 17;
            this.Save.Text = "შენახვა";
            this.Save.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Save.Textcolor = System.Drawing.Color.White;
            this.Save.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Save.Click += new System.EventHandler(this.Save_Click);
            // 
            // ChangeBox
            // 
            this.ChangeBox.BackColor = System.Drawing.SystemColors.ControlLight;
            this.ChangeBox.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.ChangeBox.Font = new System.Drawing.Font("Century Gothic", 9.75F);
            this.ChangeBox.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.ChangeBox.HintForeColor = System.Drawing.Color.Empty;
            this.ChangeBox.HintText = "ხურდა";
            this.ChangeBox.isPassword = false;
            this.ChangeBox.LineFocusedColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(100)))), ((int)(((byte)(15)))));
            this.ChangeBox.LineIdleColor = System.Drawing.Color.Gray;
            this.ChangeBox.LineMouseHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(100)))), ((int)(((byte)(15)))));
            this.ChangeBox.LineThickness = 3;
            this.ChangeBox.Location = new System.Drawing.Point(648, 212);
            this.ChangeBox.Margin = new System.Windows.Forms.Padding(4);
            this.ChangeBox.Name = "ChangeBox";
            this.ChangeBox.Size = new System.Drawing.Size(217, 44);
            this.ChangeBox.TabIndex = 18;
            this.ChangeBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Left;
            this.ChangeBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ChangeBox_KeyDown);
            // 
            // DescriptionBox
            // 
            this.DescriptionBox.BackColor = System.Drawing.SystemColors.ControlLight;
            this.DescriptionBox.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.DescriptionBox.Font = new System.Drawing.Font("Century Gothic", 9.75F);
            this.DescriptionBox.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.DescriptionBox.HintForeColor = System.Drawing.Color.Empty;
            this.DescriptionBox.HintText = "აღწერა";
            this.DescriptionBox.isPassword = false;
            this.DescriptionBox.LineFocusedColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(100)))), ((int)(((byte)(15)))));
            this.DescriptionBox.LineIdleColor = System.Drawing.Color.Gray;
            this.DescriptionBox.LineMouseHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(100)))), ((int)(((byte)(15)))));
            this.DescriptionBox.LineThickness = 3;
            this.DescriptionBox.Location = new System.Drawing.Point(648, 278);
            this.DescriptionBox.Margin = new System.Windows.Forms.Padding(4);
            this.DescriptionBox.Name = "DescriptionBox";
            this.DescriptionBox.Size = new System.Drawing.Size(419, 44);
            this.DescriptionBox.TabIndex = 19;
            this.DescriptionBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Left;
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
            this.BackBTN.Location = new System.Drawing.Point(1140, 443);
            this.BackBTN.Name = "BackBTN";
            this.BackBTN.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.BackBTN.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.BackBTN.OnHoverTextColor = System.Drawing.Color.White;
            this.BackBTN.selected = false;
            this.BackBTN.Size = new System.Drawing.Size(49, 48);
            this.BackBTN.TabIndex = 21;
            this.BackBTN.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.BackBTN.Textcolor = System.Drawing.Color.White;
            this.BackBTN.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BackBTN.Click += new System.EventHandler(this.BackBTN_Click);
            // 
            // WaiterWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.Controls.Add(this.BackBTN);
            this.Controls.Add(this.DescriptionBox);
            this.Controls.Add(this.ChangeBox);
            this.Controls.Add(this.Save);
            this.Controls.Add(this.TableClearBTN);
            this.Controls.Add(this.PayTypeComboBox);
            this.Controls.Add(this.ResultCombobox);
            this.Controls.Add(this.TypeBox);
            this.Controls.Add(this.Table_8);
            this.Controls.Add(this.Table_7);
            this.Controls.Add(this.Table_6);
            this.Controls.Add(this.Table_5);
            this.Controls.Add(this.Table_4);
            this.Controls.Add(this.Table_3);
            this.Controls.Add(this.Table_2);
            this.Controls.Add(this.Table_1);
            this.Controls.Add(this.sumOfOrder);
            this.Controls.Add(this.RemoveClick);
            this.Controls.Add(this.Delete);
            this.Controls.Add(this.dataGridView3);
            this.Controls.Add(this.dataGridView1);
            this.Name = "WaiterWindow";
            this.Size = new System.Drawing.Size(1192, 494);
            this.Load += new System.EventHandler(this.WaiterWindow_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private Bunifu.Framework.UI.BunifuCustomDataGrid dataGridView1;
        private Bunifu.Framework.UI.BunifuCustomDataGrid dataGridView3;
        private Bunifu.Framework.UI.BunifuFlatButton Delete;
        private Bunifu.Framework.UI.BunifuFlatButton RemoveClick;
        private Bunifu.Framework.UI.BunifuFlatButton sumOfOrder;
        private Bunifu.Framework.UI.BunifuFlatButton Table_1;
        private Bunifu.Framework.UI.BunifuFlatButton Table_2;
        private Bunifu.Framework.UI.BunifuFlatButton Table_3;
        private Bunifu.Framework.UI.BunifuFlatButton Table_4;
        private Bunifu.Framework.UI.BunifuFlatButton Table_5;
        private Bunifu.Framework.UI.BunifuFlatButton Table_6;
        private Bunifu.Framework.UI.BunifuFlatButton Table_7;
        private Bunifu.Framework.UI.BunifuFlatButton Table_8;
        private Bunifu.Framework.UI.BunifuDropdown TypeBox;
        private Bunifu.Framework.UI.BunifuDropdown ResultCombobox;
        private Bunifu.Framework.UI.BunifuDropdown PayTypeComboBox;
        private Bunifu.Framework.UI.BunifuFlatButton TableClearBTN;
        private Bunifu.Framework.UI.BunifuFlatButton Save;
        private Bunifu.Framework.UI.BunifuMaterialTextbox ChangeBox;
        private Bunifu.Framework.UI.BunifuMaterialTextbox DescriptionBox;
        private Bunifu.Framework.UI.BunifuFlatButton BackBTN;
    }
}
