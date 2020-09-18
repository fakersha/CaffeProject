namespace CafeProject
{
    partial class UsersRemover
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(UsersRemover));
            this.BaseDataGridView = new Bunifu.Framework.UI.BunifuCustomDataGrid();
            this.BackBTN2 = new Bunifu.Framework.UI.BunifuFlatButton();
            this.button1 = new Bunifu.Framework.UI.BunifuFlatButton();
            this.BackBTN = new Bunifu.Framework.UI.BunifuFlatButton();
            ((System.ComponentModel.ISupportInitialize)(this.BaseDataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // BaseDataGridView
            // 
            this.BaseDataGridView.AllowUserToDeleteRows = false;
            this.BaseDataGridView.AllowUserToResizeColumns = false;
            this.BaseDataGridView.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.BaseDataGridView.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.BaseDataGridView.BackgroundColor = System.Drawing.Color.Gainsboro;
            this.BaseDataGridView.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.BaseDataGridView.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Century Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.BaseDataGridView.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.BaseDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.BaseDataGridView.DoubleBuffered = true;
            this.BaseDataGridView.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.BaseDataGridView.EnableHeadersVisualStyles = false;
            this.BaseDataGridView.HeaderBgColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.BaseDataGridView.HeaderForeColor = System.Drawing.Color.White;
            this.BaseDataGridView.Location = new System.Drawing.Point(22, 66);
            this.BaseDataGridView.Name = "BaseDataGridView";
            this.BaseDataGridView.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            this.BaseDataGridView.Size = new System.Drawing.Size(541, 163);
            this.BaseDataGridView.TabIndex = 1;
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
            this.BackBTN2.TabIndex = 24;
            this.BackBTN2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.BackBTN2.Textcolor = System.Drawing.Color.White;
            this.BackBTN2.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BackBTN2.Click += new System.EventHandler(this.BackBTN2_Click);
            // 
            // button1
            // 
            this.button1.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.button1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.button1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button1.BorderRadius = 0;
            this.button1.ButtonText = "წაშლა";
            this.button1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.button1.DisabledColor = System.Drawing.Color.Gray;
            this.button1.Iconcolor = System.Drawing.Color.Transparent;
            this.button1.Iconimage = ((System.Drawing.Image)(resources.GetObject("button1.Iconimage")));
            this.button1.Iconimage_right = null;
            this.button1.Iconimage_right_Selected = null;
            this.button1.Iconimage_Selected = null;
            this.button1.IconMarginLeft = 0;
            this.button1.IconMarginRight = 0;
            this.button1.IconRightVisible = true;
            this.button1.IconRightZoom = 0D;
            this.button1.IconVisible = true;
            this.button1.IconZoom = 50D;
            this.button1.IsTab = false;
            this.button1.Location = new System.Drawing.Point(22, 252);
            this.button1.Name = "button1";
            this.button1.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.button1.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.button1.OnHoverTextColor = System.Drawing.Color.White;
            this.button1.selected = false;
            this.button1.Size = new System.Drawing.Size(170, 36);
            this.button1.TabIndex = 25;
            this.button1.Text = "წაშლა";
            this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button1.Textcolor = System.Drawing.Color.White;
            this.button1.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Click += new System.EventHandler(this.button1_Click);
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
            this.BackBTN.TabIndex = 26;
            this.BackBTN.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.BackBTN.Textcolor = System.Drawing.Color.White;
            this.BackBTN.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BackBTN.Click += new System.EventHandler(this.BackBTN_Click);
            // 
            // UsersRemover
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.Controls.Add(this.BackBTN);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.BackBTN2);
            this.Controls.Add(this.BaseDataGridView);
            this.Name = "UsersRemover";
            this.Size = new System.Drawing.Size(595, 330);
            this.Load += new System.EventHandler(this.UsersRemover_Load);
            ((System.ComponentModel.ISupportInitialize)(this.BaseDataGridView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Bunifu.Framework.UI.BunifuCustomDataGrid BaseDataGridView;
        private Bunifu.Framework.UI.BunifuFlatButton BackBTN2;
        private Bunifu.Framework.UI.BunifuFlatButton button1;
        private Bunifu.Framework.UI.BunifuFlatButton BackBTN;
    }
}
