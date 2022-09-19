using LinqToExcel;
using MetroSet_UI.Forms;
using ReConciler.Data;
using ReConciler.Model;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReConciler
{
    public partial class frmMergeSheets : MetroSetForm
    {
        public frmMergeSheets()
        {
            InitializeComponent();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"D:\",
                Title = "Browse Text Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "xls",
                Filter = "xls files (*.xlsx)|*.xlsx",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtStmtFileLoc.Text = openFileDialog1.FileName;
                btnLoadFile.Visible = true;
            }
        }

        private void btnLoadFile_Click(object sender, EventArgs e)
        {
            var excel = new ExcelQueryFactory(txtStmtFileLoc.Text.Trim());
            excel.ReadOnly = true;
            var sheets = excel.GetWorksheetNames();
            foreach (var sheet in sheets)
            {
                var sheetData = (from x in excel.Worksheet(sheet) select x).ToList();
            }

            //Create a Workbook object
            //Workbook workbook = new Workbook();
            ////Load an Excel file
            //workbook.LoadFromFile(txtStmtFileLoc.Text.Trim());

            ////Get the first worksheet
            //Worksheet sheet1 = workbook.Worksheets[0];
            ////Get the second worksheet
            //Worksheet sheet2 = workbook.Worksheets[1];

            ////Get the used range in the second worksheet
            //CellRange sourceRange = sheet2.AllocatedRange;
            ////Specify the destination range in the first worksheet
            //CellRange destRange = sheet1.Range[sheet1.LastRow + 1, 1];

            ////Copy the used range of the second worksheet to the destination range in the first worksheet
            //sourceRange.Copy(destRange);

            ////Remove the second worksheet
            //sheet2.Remove();

            ////Save the result file
            //workbook.SaveToFile(@"D:\statements\MergeWorksheets.xlsx", ExcelVersion.Version2013);
        }
    }
}
