
namespace ReConciler
{
    partial class frmConvertStatement
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmConvertStatement));
            this.metroPanel2 = new MetroFramework.Controls.MetroPanel();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.progressBar1 = new MetroFramework.Controls.MetroProgressBar();
            this.metroSetLabel1 = new MetroSet_UI.Controls.MetroSetLabel();
            this.txtStmtFileLoc = new MetroSet_UI.Controls.MetroSetTextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.metroSetLabel2 = new MetroSet_UI.Controls.MetroSetLabel();
            this.txtFolderLoc = new MetroSet_UI.Controls.MetroSetTextBox();
            this.btnBrowse = new MetroSet_UI.Controls.MetroSetButton();
            this.btnConvert = new MetroSet_UI.Controls.MetroSetButton();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.metroSetDivider1 = new MetroSet_UI.Controls.MetroSetDivider();
            this.metroSetControlBox1 = new MetroSet_UI.Controls.MetroSetControlBox();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.msBtnBrowse = new MetroSet_UI.Controls.MetroSetButton();
            //this.axAcroPDF1 = new AxAcroPDFLib.AxAcroPDF();
            this.webBrowser1 = new System.Windows.Forms.WebBrowser();
            this.button1 = new System.Windows.Forms.Button();
            this.metroPanel2.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.panel1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            //((System.ComponentModel.ISupportInitialize)(this.axAcroPDF1)).BeginInit();
            this.SuspendLayout();
            // 
            // metroPanel2
            // 
            this.metroPanel2.Controls.Add(this.webBrowser1);
            //this.metroPanel2.Controls.Add(this.axAcroPDF1);
            this.metroPanel2.HorizontalScrollbarBarColor = true;
            this.metroPanel2.HorizontalScrollbarHighlightOnWheel = false;
            this.metroPanel2.HorizontalScrollbarSize = 10;
            this.metroPanel2.Location = new System.Drawing.Point(6, 16);
            this.metroPanel2.Name = "metroPanel2";
            this.metroPanel2.Size = new System.Drawing.Size(1224, 400);
            this.metroPanel2.TabIndex = 8;
            this.metroPanel2.VerticalScrollbarBarColor = true;
            this.metroPanel2.VerticalScrollbarHighlightOnWheel = false;
            this.metroPanel2.VerticalScrollbarSize = 10;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.metroPanel2);
            this.groupBox2.Location = new System.Drawing.Point(4, 155);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(1236, 422);
            this.groupBox2.TabIndex = 7;
            this.groupBox2.TabStop = false;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(1112, 46);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(113, 23);
            this.progressBar1.TabIndex = 9;
            this.progressBar1.Visible = false;
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
            this.metroSetLabel1.Text = "Select PDF Statement:*";
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
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.groupBox2);
            this.panel1.Controls.Add(this.groupBox1);
            this.panel1.Location = new System.Drawing.Point(15, 84);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1245, 582);
            this.panel1.TabIndex = 11;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.metroSetLabel2);
            this.groupBox1.Controls.Add(this.txtFolderLoc);
            this.groupBox1.Controls.Add(this.msBtnBrowse);
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Controls.Add(this.progressBar1);
            this.groupBox1.Controls.Add(this.metroSetLabel1);
            this.groupBox1.Controls.Add(this.txtStmtFileLoc);
            this.groupBox1.Controls.Add(this.btnBrowse);
            this.groupBox1.Controls.Add(this.btnConvert);
            this.groupBox1.Location = new System.Drawing.Point(3, 2);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox1.Size = new System.Drawing.Size(1237, 148);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            // 
            // metroSetLabel2
            // 
            this.metroSetLabel2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.metroSetLabel2.IsDerivedStyle = true;
            this.metroSetLabel2.Location = new System.Drawing.Point(15, 76);
            this.metroSetLabel2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.metroSetLabel2.Name = "metroSetLabel2";
            this.metroSetLabel2.Size = new System.Drawing.Size(175, 24);
            this.metroSetLabel2.Style = MetroSet_UI.Enums.Style.Light;
            this.metroSetLabel2.StyleManager = null;
            this.metroSetLabel2.TabIndex = 13;
            this.metroSetLabel2.Text = "Converter Storage Location:";
            this.metroSetLabel2.ThemeAuthor = "Narwin";
            this.metroSetLabel2.ThemeName = "MetroLite";
            // 
            // txtFolderLoc
            // 
            this.txtFolderLoc.AutoCompleteCustomSource = null;
            this.txtFolderLoc.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.None;
            this.txtFolderLoc.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.None;
            this.txtFolderLoc.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(155)))), ((int)(((byte)(155)))), ((int)(((byte)(155)))));
            this.txtFolderLoc.DisabledBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(204)))), ((int)(((byte)(204)))), ((int)(((byte)(204)))));
            this.txtFolderLoc.DisabledBorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(155)))), ((int)(((byte)(155)))), ((int)(((byte)(155)))));
            this.txtFolderLoc.DisabledForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(136)))), ((int)(((byte)(136)))), ((int)(((byte)(136)))));
            this.txtFolderLoc.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.txtFolderLoc.HoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(102)))), ((int)(((byte)(102)))), ((int)(((byte)(102)))));
            this.txtFolderLoc.Image = null;
            this.txtFolderLoc.IsDerivedStyle = true;
            this.txtFolderLoc.Lines = null;
            this.txtFolderLoc.Location = new System.Drawing.Point(15, 102);
            this.txtFolderLoc.Margin = new System.Windows.Forms.Padding(2);
            this.txtFolderLoc.MaxLength = 32767;
            this.txtFolderLoc.Multiline = false;
            this.txtFolderLoc.Name = "txtFolderLoc";
            this.txtFolderLoc.ReadOnly = true;
            this.txtFolderLoc.Size = new System.Drawing.Size(719, 30);
            this.txtFolderLoc.Style = MetroSet_UI.Enums.Style.Light;
            this.txtFolderLoc.StyleManager = null;
            this.txtFolderLoc.TabIndex = 11;
            this.txtFolderLoc.TextAlign = System.Windows.Forms.HorizontalAlignment.Left;
            this.txtFolderLoc.ThemeAuthor = "Narwin";
            this.txtFolderLoc.ThemeName = "MetroLite";
            this.txtFolderLoc.UseSystemPasswordChar = false;
            this.txtFolderLoc.WatermarkText = "";
            // 
            // btnBrowse
            // 
            this.btnBrowse.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnBrowse.DisabledBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(120)))), ((int)(((byte)(65)))), ((int)(((byte)(177)))), ((int)(((byte)(225)))));
            this.btnBrowse.DisabledBorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(120)))), ((int)(((byte)(65)))), ((int)(((byte)(177)))), ((int)(((byte)(225)))));
            this.btnBrowse.DisabledForeColor = System.Drawing.Color.Gray;
            this.btnBrowse.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.btnBrowse.HoverBorderColor = System.Drawing.Color.Black;
            this.btnBrowse.HoverColor = System.Drawing.Color.Black;
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
            // btnConvert
            // 
            this.btnConvert.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnConvert.DisabledBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(120)))), ((int)(((byte)(65)))), ((int)(((byte)(177)))), ((int)(((byte)(225)))));
            this.btnConvert.DisabledBorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(120)))), ((int)(((byte)(65)))), ((int)(((byte)(177)))), ((int)(((byte)(225)))));
            this.btnConvert.DisabledForeColor = System.Drawing.Color.Gray;
            this.btnConvert.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.btnConvert.HoverBorderColor = System.Drawing.Color.DarkRed;
            this.btnConvert.HoverColor = System.Drawing.Color.DarkRed;
            this.btnConvert.HoverTextColor = System.Drawing.Color.White;
            this.btnConvert.IsDerivedStyle = true;
            this.btnConvert.Location = new System.Drawing.Point(808, 39);
            this.btnConvert.Margin = new System.Windows.Forms.Padding(2);
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.NormalBorderColor = System.Drawing.Color.Black;
            this.btnConvert.NormalColor = System.Drawing.Color.Black;
            this.btnConvert.NormalTextColor = System.Drawing.Color.White;
            this.btnConvert.PressBorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(35)))), ((int)(((byte)(147)))), ((int)(((byte)(195)))));
            this.btnConvert.PressColor = System.Drawing.Color.FromArgb(((int)(((byte)(35)))), ((int)(((byte)(147)))), ((int)(((byte)(195)))));
            this.btnConvert.PressTextColor = System.Drawing.Color.White;
            this.btnConvert.Size = new System.Drawing.Size(113, 30);
            this.btnConvert.Style = MetroSet_UI.Enums.Style.Custom;
            this.btnConvert.StyleManager = null;
            this.btnConvert.TabIndex = 3;
            this.btnConvert.Text = "Convert to Excel";
            this.btnConvert.ThemeAuthor = "Narwin";
            this.btnConvert.ThemeName = "MetroLite";
            this.btnConvert.Visible = false;
            this.btnConvert.Click += new System.EventHandler(this.btnConvert_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // metroSetDivider1
            // 
            this.metroSetDivider1.IsDerivedStyle = false;
            this.metroSetDivider1.Location = new System.Drawing.Point(15, 75);
            this.metroSetDivider1.Margin = new System.Windows.Forms.Padding(2);
            this.metroSetDivider1.Name = "metroSetDivider1";
            this.metroSetDivider1.Orientation = MetroSet_UI.Enums.DividerStyle.Horizontal;
            this.metroSetDivider1.Size = new System.Drawing.Size(1245, 4);
            this.metroSetDivider1.Style = MetroSet_UI.Enums.Style.Light;
            this.metroSetDivider1.StyleManager = null;
            this.metroSetDivider1.TabIndex = 10;
            this.metroSetDivider1.Text = "metroSetDivider1";
            this.metroSetDivider1.ThemeAuthor = "Narwin";
            this.metroSetDivider1.ThemeName = "MetroLite";
            this.metroSetDivider1.Thickness = 4;
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
            this.metroSetControlBox1.TabIndex = 9;
            this.metroSetControlBox1.Text = "metroSetControlBox1";
            this.metroSetControlBox1.ThemeAuthor = "Narwin";
            this.metroSetControlBox1.ThemeName = "MetroLite";
            // 
            // msBtnBrowse
            // 
            this.msBtnBrowse.Cursor = System.Windows.Forms.Cursors.Hand;
            this.msBtnBrowse.DisabledBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(120)))), ((int)(((byte)(65)))), ((int)(((byte)(177)))), ((int)(((byte)(225)))));
            this.msBtnBrowse.DisabledBorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(120)))), ((int)(((byte)(65)))), ((int)(((byte)(177)))), ((int)(((byte)(225)))));
            this.msBtnBrowse.DisabledForeColor = System.Drawing.Color.Gray;
            this.msBtnBrowse.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.msBtnBrowse.HoverBorderColor = System.Drawing.Color.Black;
            this.msBtnBrowse.HoverColor = System.Drawing.Color.Black;
            this.msBtnBrowse.HoverTextColor = System.Drawing.Color.White;
            this.msBtnBrowse.IsDerivedStyle = true;
            this.msBtnBrowse.Location = new System.Drawing.Point(734, 102);
            this.msBtnBrowse.Margin = new System.Windows.Forms.Padding(2);
            this.msBtnBrowse.Name = "msBtnBrowse";
            this.msBtnBrowse.NormalBorderColor = System.Drawing.Color.Maroon;
            this.msBtnBrowse.NormalColor = System.Drawing.Color.Maroon;
            this.msBtnBrowse.NormalTextColor = System.Drawing.Color.White;
            this.msBtnBrowse.PressBorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(35)))), ((int)(((byte)(147)))), ((int)(((byte)(195)))));
            this.msBtnBrowse.PressColor = System.Drawing.Color.FromArgb(((int)(((byte)(35)))), ((int)(((byte)(147)))), ((int)(((byte)(195)))));
            this.msBtnBrowse.PressTextColor = System.Drawing.Color.White;
            this.msBtnBrowse.Size = new System.Drawing.Size(70, 30);
            this.msBtnBrowse.Style = MetroSet_UI.Enums.Style.Custom;
            this.msBtnBrowse.StyleManager = null;
            this.msBtnBrowse.TabIndex = 12;
            this.msBtnBrowse.Text = "Browse";
            this.msBtnBrowse.ThemeAuthor = "Narwin";
            this.msBtnBrowse.ThemeName = "MetroLite";
            this.msBtnBrowse.Visible = false;
            this.msBtnBrowse.Click += new System.EventHandler(this.msBtnBrowse_Click);
            // 
            // axAcroPDF1
            // 
            //this.axAcroPDF1.Enabled = true;
            //this.axAcroPDF1.Location = new System.Drawing.Point(0, 0);
            //this.axAcroPDF1.Name = "axAcroPDF1";
            //this.axAcroPDF1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axAcroPDF1.OcxState")));
            //this.axAcroPDF1.Size = new System.Drawing.Size(1224, 400);
            //this.axAcroPDF1.TabIndex = 2;
            //this.axAcroPDF1.Enter += new System.EventHandler(this.axAcroPDF1_Enter);
            // 
            // webBrowser1
            // 
            this.webBrowser1.Location = new System.Drawing.Point(30, 23);
            this.webBrowser1.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser1.Name = "webBrowser1";
            this.webBrowser1.Size = new System.Drawing.Size(1156, 342);
            this.webBrowser1.TabIndex = 3;
            this.webBrowser1.Visible = false;
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.White;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Image = global::ReConciler.Properties.Resources._1379793_document_excel_file_spreadsheet_table_icon;
            this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button1.Location = new System.Drawing.Point(938, 39);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(168, 30);
            this.button1.TabIndex = 10;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Visible = false;
            // 
            // frmConvertStatement
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1278, 681);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.metroSetDivider1);
            this.Controls.Add(this.metroSetControlBox1);
            this.Name = "frmConvertStatement";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Convert Statement";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmConvertStatement_FormClosed);
            this.metroPanel2.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            //((System.ComponentModel.ISupportInitialize)(this.axAcroPDF1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private MetroFramework.Controls.MetroPanel metroPanel2;
        private System.Windows.Forms.GroupBox groupBox2;
        private MetroFramework.Controls.MetroProgressBar progressBar1;
        private MetroSet_UI.Controls.MetroSetLabel metroSetLabel1;
        private MetroSet_UI.Controls.MetroSetTextBox txtStmtFileLoc;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.GroupBox groupBox1;
        private MetroSet_UI.Controls.MetroSetButton btnBrowse;
        private MetroSet_UI.Controls.MetroSetButton btnConvert;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private MetroSet_UI.Controls.MetroSetDivider metroSetDivider1;
        private MetroSet_UI.Controls.MetroSetControlBox metroSetControlBox1;
        //private AxAcroPDFLib.AxAcroPDF axAcroPDF1;
        private System.Windows.Forms.Button button1;
        private MetroSet_UI.Controls.MetroSetTextBox txtFolderLoc;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private MetroSet_UI.Controls.MetroSetLabel metroSetLabel2;
        private MetroSet_UI.Controls.MetroSetButton msBtnBrowse;
        private System.Windows.Forms.WebBrowser webBrowser1;
    }
}