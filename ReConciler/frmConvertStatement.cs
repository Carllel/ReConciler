using MetroSet_UI.Forms;
using PdfSharp.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReConciler
{
    public partial class frmConvertStatement : MetroSetForm
    {
        private readonly XGraphics gfx;
        private string openFileName, folderName;
        private bool fileOpened = false;
        public frmConvertStatement()
        {
            InitializeComponent();

        }


        private void btnBrowse_Click(object sender, EventArgs e)
        {
            btnConvert.Visible = false;
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"D:\",
                Title = "Browse Text Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "pdf",
                Filter = "pdf files (*.pdf)|*.pdf",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtStmtFileLoc.Text = openFileDialog1.FileName;
                try
                {
                    //Load pdf
                    //axAcroPDF1.src = openFileDialog1.FileName;
                    //txtFolderLoc.Text = Path.GetDirectoryName(txtStmtFileLoc.Text);
                    
                    webBrowser1.Navigate(new Uri(openFileDialog1.FileName));
                    btnConvert.Visible = true;
                }
                catch (Exception ex)
                {

                    MetroSetMessageBox.Show(this, $"{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private async void btnConvert_Click(object sender, EventArgs e)
        {
            (Application.OpenForms["Main"] as Main).slblStatus.Text = "PDF conversion in progress...";
            (Application.OpenForms["Main"] as Main).pbLoader.Visible=true;
            btnConvert.Enabled = false;
            string parentFolder = Path.GetDirectoryName(txtStmtFileLoc.Text);
            string savePath = parentFolder + @"\xslxconverted\";
            txtFolderLoc.Text = savePath;
            string xlxsFileName = Path.GetFileName(txtStmtFileLoc.Text).Replace("pdf", "xlsx").Replace("PDF", "xlsx");
            if (!Directory.Exists(savePath))
            {
                Directory.CreateDirectory(savePath);
            }
            if (File.Exists((savePath + xlxsFileName)))
            {
                File.Delete((savePath + xlxsFileName));
            }
            try
            {
                //PDF2Excel.Initialize("");
                //PDF2Excel.Format format = PDF2Excel.Format.DSV;
                //format = PDF2Excel.Format.MSExcel2007;//convert to xlxs
                //PDF2Excel.Document doc = new PDF2Excel.Document(txtStmtFileLoc.Text, "", "");
                //doc.Convert((savePath + xlxsFileName), format, "");
                //doc.Close();
                //PDF2Excel.UnInitialize();
                await Task.Run(() => {
                    PDF2Excel.Initialize("");
                    PDF2Excel.Format format = PDF2Excel.Format.DSV;
                    format = PDF2Excel.Format.MSExcel2007;//convert to xlxs

                    PDF2Excel.Document doc = new PDF2Excel.Document(txtStmtFileLoc.Text, "", "");
                    doc.Convert((savePath + xlxsFileName), format, "");
                    doc.Close();
                    PDF2Excel.UnInitialize();

                });
                (Application.OpenForms["Main"] as Main).pbLoader.Visible = false;
                (Application.OpenForms["Main"] as Main).slblStatus.Text = "PDF conversion complete";
                MetroSetMessageBox.Show(this, "Conversion to excel successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (PCS.Exception exception)
            {
                (Application.OpenForms["Main"] as Main).slblStatus.Text = "PDF conversion error";
                MetroSetMessageBox.Show(this, $"{exception.GetErrorCode().ToString()}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //MessageBox.Show(exception.GetErrorCode().ToString(), "Error");
            }
            catch (System.Exception exception)
            {
                (Application.OpenForms["Main"] as Main).slblStatus.Text = "PDF conversion error";
                MetroSetMessageBox.Show(this, $"{exception.ToString()}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //MessageBox.Show(exception.ToString(), "Error");
            }

            btnConvert.Enabled = true;
            (Application.OpenForms["Main"] as Main).pbLoader.Visible = false;

        }

        private void frmConvertStatement_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.axAcroPDF1.Dispose();
            this.axAcroPDF1 = null;
        }

        private void axAcroPDF1_Enter(object sender, EventArgs e)
        {

        }

        private void msBtnBrowse_Click(object sender, EventArgs e)
        {
            // Show the FolderBrowserDialog.
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                folderName = folderBrowserDialog1.SelectedPath;
                if (!fileOpened)
                {
                    // No file is opened, bring up openFileDialog in selected path.
                    txtFolderLoc.Text = folderName;
                }
            }
        }
    }
}
