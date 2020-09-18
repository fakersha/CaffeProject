namespace CafeProject
{
    partial class LoadingWindow
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(LoadingWindow));
            this.bunifuGradientPanel1 = new Bunifu.Framework.UI.BunifuGradientPanel();
            this.LoadingLabel = new System.Windows.Forms.Label();
            this.LoadingBar = new Bunifu.Framework.UI.BunifuProgressBar();
            this.LoadingPoint = new System.Windows.Forms.Label();
            this.Procents = new System.Windows.Forms.Label();
            this.LoadingImage3 = new System.Windows.Forms.PictureBox();
            this.LoadingImage2 = new System.Windows.Forms.PictureBox();
            this.LoadingImages = new System.Windows.Forms.PictureBox();
            this.LoadingTimer = new System.Windows.Forms.Timer(this.components);
            this.SecondTimer = new System.Windows.Forms.Timer(this.components);
            this.bunifuGradientPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.LoadingImage3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.LoadingImage2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.LoadingImages)).BeginInit();
            this.SuspendLayout();
            // 
            // bunifuGradientPanel1
            // 
            this.bunifuGradientPanel1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("bunifuGradientPanel1.BackgroundImage")));
            this.bunifuGradientPanel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.bunifuGradientPanel1.Controls.Add(this.LoadingLabel);
            this.bunifuGradientPanel1.Controls.Add(this.LoadingBar);
            this.bunifuGradientPanel1.Controls.Add(this.LoadingPoint);
            this.bunifuGradientPanel1.Controls.Add(this.Procents);
            this.bunifuGradientPanel1.Controls.Add(this.LoadingImage3);
            this.bunifuGradientPanel1.Controls.Add(this.LoadingImage2);
            this.bunifuGradientPanel1.Controls.Add(this.LoadingImages);
            this.bunifuGradientPanel1.GradientBottomLeft = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.bunifuGradientPanel1.GradientBottomRight = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.bunifuGradientPanel1.GradientTopLeft = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.bunifuGradientPanel1.GradientTopRight = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(151)))), ((int)(((byte)(60)))), ((int)(((byte)(49)))));
            this.bunifuGradientPanel1.Location = new System.Drawing.Point(0, 0);
            this.bunifuGradientPanel1.Name = "bunifuGradientPanel1";
            this.bunifuGradientPanel1.Quality = 10;
            this.bunifuGradientPanel1.Size = new System.Drawing.Size(676, 476);
            this.bunifuGradientPanel1.TabIndex = 0;
            this.bunifuGradientPanel1.Paint += new System.Windows.Forms.PaintEventHandler(this.bunifuGradientPanel1_Paint);
            // 
            // LoadingLabel
            // 
            this.LoadingLabel.AutoSize = true;
            this.LoadingLabel.BackColor = System.Drawing.Color.Transparent;
            this.LoadingLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LoadingLabel.ForeColor = System.Drawing.Color.White;
            this.LoadingLabel.Location = new System.Drawing.Point(28, 431);
            this.LoadingLabel.Name = "LoadingLabel";
            this.LoadingLabel.Size = new System.Drawing.Size(132, 25);
            this.LoadingLabel.TabIndex = 20;
            this.LoadingLabel.Text = "იტვირთება";
            // 
            // LoadingBar
            // 
            this.LoadingBar.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(117)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.LoadingBar.BorderRadius = 5;
            this.LoadingBar.Location = new System.Drawing.Point(30, 390);
            this.LoadingBar.MaximumValue = 100;
            this.LoadingBar.Name = "LoadingBar";
            this.LoadingBar.ProgressColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(100)))), ((int)(((byte)(15)))));
            this.LoadingBar.Size = new System.Drawing.Size(616, 28);
            this.LoadingBar.TabIndex = 19;
            this.LoadingBar.Value = 0;
            // 
            // LoadingPoint
            // 
            this.LoadingPoint.AutoSize = true;
            this.LoadingPoint.BackColor = System.Drawing.Color.Transparent;
            this.LoadingPoint.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LoadingPoint.ForeColor = System.Drawing.Color.Transparent;
            this.LoadingPoint.Location = new System.Drawing.Point(146, 407);
            this.LoadingPoint.Name = "LoadingPoint";
            this.LoadingPoint.Size = new System.Drawing.Size(63, 55);
            this.LoadingPoint.TabIndex = 18;
            this.LoadingPoint.Text = "...";
            // 
            // Procents
            // 
            this.Procents.AutoSize = true;
            this.Procents.BackColor = System.Drawing.Color.Transparent;
            this.Procents.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Procents.ForeColor = System.Drawing.Color.White;
            this.Procents.Location = new System.Drawing.Point(599, 431);
            this.Procents.Name = "Procents";
            this.Procents.Size = new System.Drawing.Size(43, 25);
            this.Procents.TabIndex = 17;
            this.Procents.Text = "0%";
            // 
            // LoadingImage3
            // 
            this.LoadingImage3.Image = ((System.Drawing.Image)(resources.GetObject("LoadingImage3.Image")));
            this.LoadingImage3.Location = new System.Drawing.Point(30, 22);
            this.LoadingImage3.Name = "LoadingImage3";
            this.LoadingImage3.Size = new System.Drawing.Size(616, 344);
            this.LoadingImage3.TabIndex = 15;
            this.LoadingImage3.TabStop = false;
            // 
            // LoadingImage2
            // 
            this.LoadingImage2.Image = ((System.Drawing.Image)(resources.GetObject("LoadingImage2.Image")));
            this.LoadingImage2.Location = new System.Drawing.Point(30, 22);
            this.LoadingImage2.Name = "LoadingImage2";
            this.LoadingImage2.Size = new System.Drawing.Size(616, 344);
            this.LoadingImage2.TabIndex = 14;
            this.LoadingImage2.TabStop = false;
            // 
            // LoadingImages
            // 
            this.LoadingImages.Image = ((System.Drawing.Image)(resources.GetObject("LoadingImages.Image")));
            this.LoadingImages.Location = new System.Drawing.Point(30, 22);
            this.LoadingImages.Name = "LoadingImages";
            this.LoadingImages.Size = new System.Drawing.Size(616, 344);
            this.LoadingImages.TabIndex = 13;
            this.LoadingImages.TabStop = false;
            // 
            // LoadingTimer
            // 
            this.LoadingTimer.Interval = 900;
            this.LoadingTimer.Tick += new System.EventHandler(this.LoadingTimer_Tick);
            // 
            // SecondTimer
            // 
            this.SecondTimer.Interval = 550;
            this.SecondTimer.Tick += new System.EventHandler(this.SecondTimer_Tick);
            // 
            // LoadingWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.bunifuGradientPanel1);
            this.Name = "LoadingWindow";
            this.Size = new System.Drawing.Size(676, 476);
            this.Load += new System.EventHandler(this.LoadingWindow_Load);
            this.bunifuGradientPanel1.ResumeLayout(false);
            this.bunifuGradientPanel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.LoadingImage3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.LoadingImage2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.LoadingImages)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Bunifu.Framework.UI.BunifuGradientPanel bunifuGradientPanel1;
        private System.Windows.Forms.Label Procents;
        private System.Windows.Forms.PictureBox LoadingImage3;
        private System.Windows.Forms.PictureBox LoadingImage2;
        private System.Windows.Forms.PictureBox LoadingImages;
        private System.Windows.Forms.Timer LoadingTimer;
        private System.Windows.Forms.Timer SecondTimer;
        private Bunifu.Framework.UI.BunifuProgressBar LoadingBar;
        private System.Windows.Forms.Label LoadingPoint;
        private System.Windows.Forms.Label LoadingLabel;
    }
}
