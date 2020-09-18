namespace CafeProject
{
    partial class CMessageBox
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CMessageBox));
            this.MessageBoxText = new System.Windows.Forms.Label();
            this.ConfirmBTN = new Bunifu.Framework.UI.BunifuFlatButton();
            this.SuspendLayout();
            // 
            // MessageBoxText
            // 
            this.MessageBoxText.AutoSize = true;
            this.MessageBoxText.BackColor = System.Drawing.Color.Transparent;
            this.MessageBoxText.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MessageBoxText.ForeColor = System.Drawing.Color.White;
            this.MessageBoxText.Location = new System.Drawing.Point(163, 71);
            this.MessageBoxText.Name = "MessageBoxText";
            this.MessageBoxText.Size = new System.Drawing.Size(71, 20);
            this.MessageBoxText.TabIndex = 29;
            this.MessageBoxText.Text = "ტექსტი";
            this.MessageBoxText.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // ConfirmBTN
            // 
            this.ConfirmBTN.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.ConfirmBTN.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.ConfirmBTN.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ConfirmBTN.BorderRadius = 0;
            this.ConfirmBTN.ButtonText = "დადასტურება";
            this.ConfirmBTN.Cursor = System.Windows.Forms.Cursors.Hand;
            this.ConfirmBTN.DisabledColor = System.Drawing.Color.Gray;
            this.ConfirmBTN.Iconcolor = System.Drawing.Color.Transparent;
            this.ConfirmBTN.Iconimage = ((System.Drawing.Image)(resources.GetObject("ConfirmBTN.Iconimage")));
            this.ConfirmBTN.Iconimage_right = null;
            this.ConfirmBTN.Iconimage_right_Selected = null;
            this.ConfirmBTN.Iconimage_Selected = null;
            this.ConfirmBTN.IconMarginLeft = 0;
            this.ConfirmBTN.IconMarginRight = 0;
            this.ConfirmBTN.IconRightVisible = true;
            this.ConfirmBTN.IconRightZoom = 0D;
            this.ConfirmBTN.IconVisible = true;
            this.ConfirmBTN.IconZoom = 50D;
            this.ConfirmBTN.IsTab = false;
            this.ConfirmBTN.Location = new System.Drawing.Point(124, 172);
            this.ConfirmBTN.Name = "ConfirmBTN";
            this.ConfirmBTN.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.ConfirmBTN.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(2)))), ((int)(((byte)(2)))));
            this.ConfirmBTN.OnHoverTextColor = System.Drawing.Color.White;
            this.ConfirmBTN.selected = false;
            this.ConfirmBTN.Size = new System.Drawing.Size(149, 36);
            this.ConfirmBTN.TabIndex = 28;
            this.ConfirmBTN.Text = "დადასტურება";
            this.ConfirmBTN.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.ConfirmBTN.Textcolor = System.Drawing.Color.White;
            this.ConfirmBTN.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ConfirmBTN.Click += new System.EventHandler(this.ConfirmBTN_Click);
            // 
            // CMessageBox
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.ClientSize = new System.Drawing.Size(381, 233);
            this.Controls.Add(this.MessageBoxText);
            this.Controls.Add(this.ConfirmBTN);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "CMessageBox";
            this.Text = "CMessageBox";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label MessageBoxText;
        private Bunifu.Framework.UI.BunifuFlatButton ConfirmBTN;
    }
}