namespace ReConciler
{
    partial class frmMergeSheets
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
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.metroSetControlBox1 = new MetroSet_UI.Controls.MetroSetControlBox();
            this.metroSetDivider1 = new MetroSet_UI.Controls.MetroSetDivider();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.metroSetLabel1 = new MetroSet_UI.Controls.MetroSetLabel();
            this.txtStmtFileLoc = new MetroSet_UI.Controls.MetroSetTextBox();
            this.btnBrowse = new MetroSet_UI.Controls.MetroSetButton();
            this.btnLoadFile = new MetroSet_UI.Controls.MetroSetButton();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // metroSetControlBox1
            // 
            this.metroSetControlBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.metroSetControlBox1.CloseHoverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(183)))), ((int)(((byte)(40)))), ((int)(((byte)(40)))));
            this.metroSetControlBox1.CloseHoverForeColor = System.Drawing.Color.White;
            this.metroSetControlBox1.CloseNormalForeColor = System.Drawing.Color.Gray;
            this.metroSetControlBox1.DisabledForeColor = System.Drawing.Color.DimGray;
            this.metroSetControlBox1.IsDerivedStyle = true;
            this.metroSetControlBox1.Location = new System.Drawing.Point(1156, 8);
            this.metroSetControlBox1.Margin = new System.Windows.Forms.Padding(2);
            this.metroSetControlBox1.MaximizeBox = true;
            this.metroSetControlBox1.MaximizeHoverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(238)))), ((int)(((byte)(238)))));
            this.metroSetControlBox1.MaximizeHoverForeColor = System.Drawing.Color.Gray;
            this.metroSetControlBox1.MaximizeNormalForeColor = System.Drawing.Color.Gray;
            this.metroSetControlBox1.MinimizeBox = false;
            this.metroSetControlBox1.MinimizeHoverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(238)))), ((int)(((byte)(238)))));
            this.metroSetControlBox1.MinimizeHoverForeColor = System.Drawing.Color.Gray;
            this.metroSetControlBox1.MinimizeNormalForeColor = System.Drawing.Color.Gray;
            this.metroSetControlBox1.Name = "metroSetControlBox1";
            this.metroSetControlBox1.Size = new System.Drawing.Size(100, 25);
            this.metroSetControlBox1.Style = MetroSet_UI.Enums.Style.Light;
            this.metroSetControlBox1.StyleManager = null;
            this.metroSetControlBox1.TabIndex = 1;
            this.metroSetControlBox1.Text = "metroSetControlBox1";
            this.metroSetControlBox1.ThemeAuthor = "Narwin";
            this.metroSetControlBox1.ThemeName = "MetroLite";
            // 
            // metroSetDivider1
            // 
            this.metroSetDivider1.IsDerivedStyle = false;
            this.metroSetDivider1.Location = new System.Drawing.Point(12, 82);
            this.metroSetDivider1.Margin = new System.Windows.Forms.Padding(2);
            this.metroSetDivider1.Name = "metroSetDivider1";
            this.metroSetDivider1.Orientation = MetroSet_UI.Enums.DividerStyle.Horizontal;
            this.metroSetDivider1.Size = new System.Drawing.Size(1245, 4);
            this.metroSetDivider1.Style = MetroSet_UI.Enums.Style.Light;
            this.metroSetDivider1.StyleManager = null;
            this.metroSetDivider1.TabIndex = 2;
            this.metroSetDivider1.Text = "metroSetDivider1";
            this.metroSetDivider1.ThemeAuthor = "Narwin";
            this.metroSetDivider1.ThemeName = "MetroLite";
            this.metroSetDivider1.Thickness = 4;
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.Color.White;
            this.groupBox1.Controls.Add(this.metroSetLabel1);
            this.groupBox1.Controls.Add(this.txtStmtFileLoc);
            this.groupBox1.Controls.Add(this.btnBrowse);
            this.groupBox1.Controls.Add(this.btnLoadFile);
            this.groupBox1.Location = new System.Drawing.Point(17, 91);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox1.Size = new System.Drawing.Size(1237, 108);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            // 
            // metroSetLabel1
            // 
            this.metroSetLabel1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.metroSetLabel1.IsDerivedStyle = true;
            this.metroSetLabel1.Location = new System.Drawing.Point(15, 15);
            this.metroSetLabel1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.metroSetLabel1.Name = "metroSetLabel1";
            this.metroSetLabel1.Size = new System.Drawing.Size(175, 24);
            this.metroSetLabel1.Style = MetroSet_UI.Enums.Style.Light;
            this.metroSetLabel1.StyleManager = null;
            this.metroSetLabel1.TabIndex = 0;
            this.metroSetLabel1.Text = "Select Excel Workbook*";
            this.metroSetLabel1.ThemeAuthor = "Narwin";
            this.metroSetLabel1.ThemeName = "MetroLite";
            // 
            // txtStmtFileLoc
            // 
            this.txtStmtFileLoc.AutoCompleteCustomSource = null;
            this.txtStmtFileLoc.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.None;
            this.txtStmtFileLoc.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.None;
            this.txtStmtFileLoc.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(155)))), ((int)(((byte)(155)))), ((int)(((byte)(155)))));
            this.txtStmtFileLoc.DisabledBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(204)))), ((int)(((byte)(204)))), ((int)(((byte)(204)))));
            this.txtStmtFileLoc.DisabledBorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(155)))), ((int)(((byte)(155)))), ((int)(((byte)(155)))));
            this.txtStmtFileLoc.DisabledForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(136)))), ((int)(((byte)(136)))), ((int)(((byte)(136)))));
            this.txtStmtFileLoc.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.txtStmtFileLoc.HoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(102)))), ((int)(((byte)(102)))), ((int)(((byte)(102)))));
            this.txtStmtFileLoc.Image = null;
            this.txtStmtFileLoc.IsDerivedStyle = true;
            this.txtStmtFileLoc.Lines = null;
            this.txtStmtFileLoc.Location = new System.Drawing.Point(15, 39);
            this.txtStmtFileLoc.Margin = new System.Windows.Forms.Padding(2);
            this.txtStmtFileLoc.MaxLength = 32767;
            this.txtStmtFileLoc.Multiline = false;
            this.txtStmtFileLoc.Name = "txtStmtFileLoc";
            this.txtStmtFileLoc.ReadOnly = true;
            this.txtStmtFileLoc.Size = new System.Drawing.Size(719, 30);
            this.txtStmtFileLoc.Style = MetroSet_UI.Enums.Style.Light;
            this.txtStmtFileLoc.StyleManager = null;
            this.txtStmtFileLoc.TabIndex = 1;
            this.txtStmtFileLoc.TextAlign = System.Windows.Forms.HorizontalAlignment.Left;
            this.txtStmtFileLoc.ThemeAuthor = "Narwin";
            this.txtStmtFileLoc.ThemeName = "MetroLite";
            this.txtStmtFileLoc.UseSystemPasswordChar = false;
            this.txtStmtFileLoc.WatermarkText = "";
            // 
            // btnBrowse
            // 
            this.btnBrowse.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnBrowse.DisabledBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(120)))), ((int)(((byte)(65)))), ((int)(((byte)(177)))), ((int)(((byte)(225)))));
            this.btnBrowse.DisabledBorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(120)))), ((int)(((byte)(65)))), ((int)(((byte)(177)))), ((int)(((byte)(225)))));
            this.btnBrowse.DisabledForeColor = System.Drawing.Color.Gray;
            this.btnBrowse.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.btnBrowse.HoverBorderColor = System.Drawing.Color.DarkBlue;
            this.btnBrowse.HoverColor = System.Drawing.Color.DarkBlue;
            this.btnBrowse.HoverTextColor = System.Drawing.Color.White;
            this.btnBrowse.IsDerivedStyle = true;
            this.btnBrowse.Location = new System.Drawing.Point(734, 39);
            this.btnBrowse.Margin = new System.Windows.Forms.Padding(2);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.NormalBorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(65)))), ((int)(((byte)(177)))), ((int)(((byte)(225)))));
            this.btnBrowse.NormalColor = System.Drawing.Color.FromArgb(((int)(((byte)(65)))), ((int)(((byte)(177)))), ((int)(((byte)(225)))));
            this.btnBrowse.NormalTextColor = System.Drawing.Color.White;
            this.btnBrowse.PressBorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(35)))), ((int)(((byte)(147)))), ((int)(((byte)(195)))));
            this.btnBrowse.PressColor = System.Drawing.Color.FromArgb(((int)(((byte)(35)))), ((int)(((byte)(147)))), ((int)(((byte)(195)))));
            this.btnBrowse.PressTextColor = System.Drawing.Color.White;
            this.btnBrowse.Size = new System.Drawing.Size(70, 30);
            this.btnBrowse.Style = MetroSet_UI.Enums.Style.Custom;
            this.btnBrowse.StyleManager = null;
            this.btnBrowse.TabIndex = 2;
            this.btnBrowse.Text = "Browse";
            this.btnBrowse.ThemeAuthor = "Narwin";
            this.btnBrowse.ThemeName = "MetroLite";
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // btnLoadFile
            // 
            this.btnLoadFile.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnLoadFile.DisabledBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(120)))), ((int)(((byte)(65)))), ((int)(((byte)(177)))), ((int)(((byte)(225)))));
            this.btnLoadFile.DisabledBorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(120)))), ((int)(((byte)(65)))), ((int)(((byte)(177)))), ((int)(((byte)(225)))));
            this.btnLoadFile.DisabledForeColor = System.Drawing.Color.Gray;
            this.btnLoadFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.btnLoadFile.HoverBorderColor = System.Drawing.Color.DarkGreen;
            this.btnLoadFile.HoverColor = System.Drawing.Color.DarkGreen;
            this.btnLoadFile.HoverTextColor = System.Drawing.Color.White;
            this.btnLoadFile.IsDerivedStyle = true;
            this.btnLoadFile.Location = new System.Drawing.Point(808, 39);
            this.btnLoadFile.Margin = new System.Windows.Forms.Padding(2);
            this.btnLoadFile.Name = "btnLoadFile";
            this.btnLoadFile.NormalBorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(65)))), ((int)(((byte)(177)))), ((int)(((byte)(225)))));
            this.btnLoadFile.NormalColor = System.Drawing.Color.FromArgb(((int)(((byte)(65)))), ((int)(((byte)(177)))), ((int)(((byte)(225)))));
            this.btnLoadFile.NormalTextColor = System.Drawing.Color.White;
            this.btnLoadFile.PressBorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(35)))), ((int)(((byte)(147)))), ((int)(((byte)(195)))));
            this.btnLoadFile.PressColor = System.Drawing.Color.FromArgb(((int)(((byte)(35)))), ((int)(((byte)(147)))), ((int)(((byte)(195)))));
            this.btnLoadFile.PressTextColor = System.Drawing.Color.White;
            this.btnLoadFile.Size = new System.Drawing.Size(113, 30);
            this.btnLoadFile.Style = MetroSet_UI.Enums.Style.Custom;
            this.btnLoadFile.StyleManager = null;
            this.btnLoadFile.TabIndex = 3;
            this.btnLoadFile.Text = "Load Workbook";
            this.btnLoadFile.ThemeAuthor = "Narwin";
            this.btnLoadFile.ThemeName = "MetroLite";
            this.btnLoadFile.Visible = false;
            this.btnLoadFile.Click += new System.EventHandler(this.btnLoadFile_Click);
            // 
            // frmMergeSheets
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1270, 567);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.metroSetDivider1);
            this.Controls.Add(this.metroSetControlBox1);
            this.Name = "frmMergeSheets";
            this.Text = "Merge Sheets";
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private MetroSet_UI.Controls.MetroSetControlBox metroSetControlBox1;
        private MetroSet_UI.Controls.MetroSetDivider metroSetDivider1;
        private System.Windows.Forms.GroupBox groupBox1;
        private MetroSet_UI.Controls.MetroSetLabel metroSetLabel1;
        private MetroSet_UI.Controls.MetroSetTextBox txtStmtFileLoc;
        private MetroSet_UI.Controls.MetroSetButton btnBrowse;
        private MetroSet_UI.Controls.MetroSetButton btnLoadFile;
    }
}