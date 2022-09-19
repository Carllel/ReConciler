
namespace ReConciler
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
            this.metroSetControlBox1 = new MetroSet_UI.Controls.MetroSetControlBox();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripBtnLoadStmt = new System.Windows.Forms.ToolStripButton();
            this.toolStripBtnConvert = new System.Windows.Forms.ToolStripButton();
            this.toolStripBtnMerge = new System.Windows.Forms.ToolStripButton();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabel2 = new System.Windows.Forms.ToolStripStatusLabel();
            this.slblStatus = new System.Windows.Forms.ToolStripStatusLabel();
            this.pbLoader = new System.Windows.Forms.PictureBox();
            this.toolStrip1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbLoader)).BeginInit();
            this.SuspendLayout();
            // 
            // metroSetControlBox1
            // 
            this.metroSetControlBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.metroSetControlBox1.CloseHoverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(183)))), ((int)(((byte)(40)))), ((int)(((byte)(40)))));
            this.metroSetControlBox1.CloseHoverForeColor = System.Drawing.Color.White;
            this.metroSetControlBox1.CloseNormalForeColor = System.Drawing.Color.Gray;
            this.metroSetControlBox1.DisabledForeColor = System.Drawing.Color.DimGray;
            this.metroSetControlBox1.IsDerivedStyle = true;
            this.metroSetControlBox1.Location = new System.Drawing.Point(1379, 11);
            this.metroSetControlBox1.MaximizeBox = true;
            this.metroSetControlBox1.MaximizeHoverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(238)))), ((int)(((byte)(238)))));
            this.metroSetControlBox1.MaximizeHoverForeColor = System.Drawing.Color.Gray;
            this.metroSetControlBox1.MaximizeNormalForeColor = System.Drawing.Color.Gray;
            this.metroSetControlBox1.MinimizeBox = true;
            this.metroSetControlBox1.MinimizeHoverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(238)))), ((int)(((byte)(238)))));
            this.metroSetControlBox1.MinimizeHoverForeColor = System.Drawing.Color.Gray;
            this.metroSetControlBox1.MinimizeNormalForeColor = System.Drawing.Color.Gray;
            this.metroSetControlBox1.Name = "metroSetControlBox1";
            this.metroSetControlBox1.Size = new System.Drawing.Size(100, 25);
            this.metroSetControlBox1.Style = MetroSet_UI.Enums.Style.Light;
            this.metroSetControlBox1.StyleManager = null;
            this.metroSetControlBox1.TabIndex = 4;
            this.metroSetControlBox1.Text = "metroSetControlBox1";
            this.metroSetControlBox1.ThemeAuthor = "Narwin";
            this.metroSetControlBox1.ThemeName = "MetroLite";
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripBtnLoadStmt,
            this.toolStripBtnConvert,
            this.toolStripBtnMerge});
            this.toolStrip1.Location = new System.Drawing.Point(12, 70);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(1453, 25);
            this.toolStrip1.TabIndex = 5;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripBtnLoadStmt
            // 
            this.toolStripBtnLoadStmt.Image = global::ReConciler.Properties.Resources.load_upload_icon;
            this.toolStripBtnLoadStmt.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripBtnLoadStmt.Name = "toolStripBtnLoadStmt";
            this.toolStripBtnLoadStmt.Size = new System.Drawing.Size(110, 22);
            this.toolStripBtnLoadStmt.Text = "Load Statement";
            this.toolStripBtnLoadStmt.Click += new System.EventHandler(this.toolStripBtnLoadStmt_Click);
            // 
            // toolStripBtnConvert
            // 
            this.toolStripBtnConvert.Image = ((System.Drawing.Image)(resources.GetObject("toolStripBtnConvert.Image")));
            this.toolStripBtnConvert.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripBtnConvert.Name = "toolStripBtnConvert";
            this.toolStripBtnConvert.Size = new System.Drawing.Size(126, 22);
            this.toolStripBtnConvert.Text = "Convert Statement";
            this.toolStripBtnConvert.Visible = false;
            this.toolStripBtnConvert.Click += new System.EventHandler(this.toolStripBtnConvert_Click);
            // 
            // toolStripBtnMerge
            // 
            this.toolStripBtnMerge.Image = ((System.Drawing.Image)(resources.GetObject("toolStripBtnMerge.Image")));
            this.toolStripBtnMerge.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripBtnMerge.Name = "toolStripBtnMerge";
            this.toolStripBtnMerge.Size = new System.Drawing.Size(98, 22);
            this.toolStripBtnMerge.Text = "Merge Sheets";
            this.toolStripBtnMerge.Visible = false;
            this.toolStripBtnMerge.Click += new System.EventHandler(this.toolStripBtnMerge_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1,
            this.toolStripStatusLabel2,
            this.slblStatus});
            this.statusStrip1.Location = new System.Drawing.Point(12, 723);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(1453, 22);
            this.statusStrip1.TabIndex = 6;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(42, 17);
            this.toolStripStatusLabel1.Text = "Status";
            // 
            // toolStripStatusLabel2
            // 
            this.toolStripStatusLabel2.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.toolStripStatusLabel2.Name = "toolStripStatusLabel2";
            this.toolStripStatusLabel2.Size = new System.Drawing.Size(17, 17);
            this.toolStripStatusLabel2.Text = " | ";
            // 
            // slblStatus
            // 
            this.slblStatus.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.slblStatus.Name = "slblStatus";
            this.slblStatus.Size = new System.Drawing.Size(0, 17);
            // 
            // pbLoader
            // 
            this.pbLoader.Image = global::ReConciler.Properties.Resources.loadersr;
            this.pbLoader.Location = new System.Drawing.Point(1273, 725);
            this.pbLoader.Name = "pbLoader";
            this.pbLoader.Size = new System.Drawing.Size(169, 20);
            this.pbLoader.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.pbLoader.TabIndex = 8;
            this.pbLoader.TabStop = false;
            this.pbLoader.Visible = false;
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundColor = System.Drawing.Color.WhiteSmoke;
            this.ClientSize = new System.Drawing.Size(1477, 757);
            this.Controls.Add(this.pbLoader);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.metroSetControlBox1);
            this.IsMdiContainer = true;
            this.Name = "Main";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Recon Optimizer";
            this.Load += new System.EventHandler(this.Main_Load);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbLoader)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private MetroSet_UI.Controls.MetroSetControlBox metroSetControlBox1;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripButton toolStripBtnLoadStmt;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel2;
        public System.Windows.Forms.ToolStripStatusLabel slblStatus;
        private System.Windows.Forms.ToolStripButton toolStripBtnConvert;
        public System.Windows.Forms.PictureBox pbLoader;
        private System.Windows.Forms.ToolStripButton toolStripBtnMerge;
    }
}

