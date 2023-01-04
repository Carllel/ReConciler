using ClosedXML.Excel;
using Newtonsoft.Json;
using ReConciler.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReConciler.Util
{
    public static class UtilManager
    {
        public static JSONReadWrite readWrite = new JSONReadWrite();

        public static List<DocumentType> GetDocumentTypes(string fileName)
        {
            List<DocumentType> stores = new List<DocumentType>();
            try
            {
                if (File.Exists(fileName))
                {
                    JSONReadWrite readWrite = new JSONReadWrite();
                    stores = JsonConvert.DeserializeObject<List<DocumentType>>(readWrite.Read(fileName));

                }
                return stores;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public static List<Vendor> GetVendors(string fileName)
        {
            List<Vendor> vendors = new List<Vendor>();
            try
            {
                if (File.Exists(fileName))
                {
                    JSONReadWrite readWrite = new JSONReadWrite();
                    vendors = JsonConvert.DeserializeObject<List<Vendor>>(readWrite.Read(fileName)).ToList();

                }
                return vendors;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
    
        public static bool CreateReportsWorkbook(string filename,string sheetname1,string sheetname2)
        {
            bool success = false;
            try
            {
                //Create new workbook
                IXLWorkbook workbook = new XLWorkbook();
                //Adding a worksheet
                IXLWorksheet worksheet = workbook.Worksheets.Add(sheetname1);
                worksheet = workbook.Worksheets.Add(sheetname2);
                //Saving the workbook
                workbook.SaveAs(filename);

                workbook.Dispose();
                success = true;
                return success;
            }
            catch (Exception ex)
            {

                return success;
            }
        }
    
        public static bool WriteReportHeaderData(string sheetname, string fpath, string rptitle, string titlecol,string titlerow, List<string> headerNames,int headerRow)
        {
            bool success = false;


            try
            {
                XLWorkbook workbook = new XLWorkbook(fpath);
                IXLWorksheet ws = workbook.Worksheet(sheetname);
                //Clear worksheet
                ws.Clear();

                #region WRITE HEADER
                //Report Header
                ws.Cell($"{titlecol}{titlerow}").Value = rptitle;
                ws.Cell($"{titlecol}{titlerow}").Style.Fill.SetBackgroundColor(XLColor.Blue);
                // Set the color for the entire cell
                ws.Cell($"{titlecol}{titlerow}").Style.Font.FontColor = XLColor.White;
                ws.Cell($"{titlecol}{titlerow}").Style.Font.Bold = true;
                ws.Cell($"{titlecol}{titlerow}").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                //Column headings

                foreach (var txt in headerNames)
                {
                    int cntHd = headerNames.IndexOf(txt) + 2;
                    ws.Cell(headerRow, cntHd).Value = txt;
                    ws.Cell(headerRow, cntHd).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    ws.Cell(headerRow, cntHd).Style.Fill.SetBackgroundColor(XLColor.Blue);
                    // Set the color for the entire cell
                    ws.Cell(headerRow, cntHd).Style.Font.FontColor = XLColor.White;
                    ws.Cell(headerRow, cntHd).Style.Font.Bold = true;
                }
                #endregion

                workbook.SaveAs(fpath);
                success = true;

                return success;
            }
            catch (Exception ex)
            {
                success = false;
            }
            return success;
        }

        public static bool WriteReportDetailsTGEDDES(string sheetname, string fpath, int startrow,DataGridView dataGrid,bool ismatched)
        {
            bool success = false;
            decimal totAmt = 0, totBal = 0;
            int lrow = 0, lrow2 = 0; 
            XLWorkbook workbook = new XLWorkbook(fpath);

            IXLWorksheet ws = workbook.Worksheet(sheetname);

            try
            {
                #region WRITE DETAILS
                foreach (DataGridViewRow r in dataGrid.Rows)
                {
                    //Look for unmatched statements
                    if (Convert.ToBoolean(r.Cells["chbIsMatch"].Value) == ismatched)
                    {
                        //Write data to column B
                        //ws.Cell($"B{startrow}").Value = r.Cells["ReferenceNo"].Value.ToString();
                        ws.Cell($"B{startrow}").Value = r.Cells[dataGrid.Columns[1].HeaderText].Value.ToString();
                        //Write data to column C
                        ws.Cell($"C{startrow}").Value = r.Cells[dataGrid.Columns[2].HeaderText].Value.ToString();
                        //Write data to column D
                        ws.Cell($"D{startrow}").Value = r.Cells[dataGrid.Columns[3].HeaderText].Value.ToString();
                        //Write data to column E
                        ws.Cell($"E{startrow}").Value = r.Cells[dataGrid.Columns[4].HeaderText].Value.ToString().Replace(",", "");
                        ws.Cell($"E{startrow}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
                        totAmt = totAmt + Convert.ToDecimal(r.Cells[dataGrid.Columns[4].HeaderText].Value.ToString().Replace(",", ""));
                        //Write data to column F
                        ws.Cell($"F{startrow}").Value = r.Cells[dataGrid.Columns[5].HeaderText].Value.ToString().Replace(",", "");
                        ws.Cell($"F{startrow}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
                        totBal = totBal + Convert.ToDecimal(r.Cells[dataGrid.Columns[5].HeaderText].Value.ToString().Replace(",", ""));

                        startrow += 1;
                    }
                }
                #endregion

                #region STYLE REPORT
                //Styling
                lrow = ws.LastRowUsed().RowNumber();
                ws.Cell($"E{lrow + 1}").Value = totAmt.ToString();
                ws.Cell($"E{lrow + 1}").Style.Font.Bold = true;
                ws.Cell($"E{lrow + 1}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
                ws.Cell($"F{lrow + 1}").Value = totBal.ToString();
                ws.Cell($"F{lrow + 1}").Style.Font.Bold = true;
                ws.Cell($"F{lrow + 1}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";

                int newRow = lrow + 1;
                ws.Range($"B{newRow}:F{newRow}").Style.Border.TopBorder = XLBorderStyleValues.Thin;

                lrow2 = ws.LastRowUsed().RowNumber();

                IXLRange range = ws.Range(ws.Cell($"B{startrow}").Address, ws.Cell($"F{lrow2}").Address);

                range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                ws.Columns().AdjustToContents();
                #endregion

                workbook.SaveAs(fpath);
                success = true;
            }
            catch (Exception ex)
            {

            }

            return success;
        }

        public static bool WriteReportDetailsCOPPERWOOD(string sheetname, string fpath, int startrow, DataGridView dataGrid, bool ismatched)
        {
            bool success = false;
            decimal totAmt = 0, totBal = 0;
            int lrow = 0, lrow2 = 0;
            XLWorkbook workbook = new XLWorkbook(fpath);

            IXLWorksheet ws = workbook.Worksheet(sheetname);

            try
            {
                #region WRITE DETAILS
                foreach (DataGridViewRow r in dataGrid.Rows)
                {
                    //Look for unmatched statements
                    if (Convert.ToBoolean(r.Cells["chbIsMatch"].Value) == ismatched)
                    {
                        //Write data to column B
                        //ws.Cell($"B{startrow}").Value = r.Cells["ReferenceNo"].Value.ToString();
                        ws.Cell($"B{startrow}").Value = r.Cells[dataGrid.Columns[1].HeaderText].Value.ToString();
                        //Write data to column C
                        ws.Cell($"C{startrow}").Value = r.Cells[dataGrid.Columns[2].HeaderText].Value.ToString();
                        //Write data to column D
                        ws.Cell($"D{startrow}").Value = r.Cells[dataGrid.Columns[3].HeaderText].Value.ToString();
                        //Write data to column E
                        ws.Cell($"E{startrow}").Value = r.Cells[dataGrid.Columns[4].HeaderText].Value.ToString().Replace(",", "");
                        ws.Cell($"E{startrow}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
                        totAmt = totAmt + Convert.ToDecimal(r.Cells[dataGrid.Columns[4].HeaderText].Value.ToString().Replace(",", ""));
                        //Write data to column F
                        ws.Cell($"F{startrow}").Value = r.Cells[dataGrid.Columns[5].HeaderText].Value.ToString().Replace(",", "");
                        ws.Cell($"F{startrow}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
                        totBal = totBal + Convert.ToDecimal(r.Cells[dataGrid.Columns[5].HeaderText].Value.ToString().Replace(",", ""));

                        startrow += 1;
                    }
                }
                #endregion

                #region STYLE REPORT
                //Styling
                lrow = ws.LastRowUsed().RowNumber();
                ws.Cell($"E{lrow + 1}").Value = totAmt.ToString();
                ws.Cell($"E{lrow + 1}").Style.Font.Bold = true;
                ws.Cell($"E{lrow + 1}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
                ws.Cell($"F{lrow + 1}").Value = totBal.ToString();
                ws.Cell($"F{lrow + 1}").Style.Font.Bold = true;
                ws.Cell($"F{lrow + 1}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";

                int newRow = lrow + 1;
                ws.Range($"B{newRow}:F{newRow}").Style.Border.TopBorder = XLBorderStyleValues.Thin;

                lrow2 = ws.LastRowUsed().RowNumber();

                IXLRange range = ws.Range(ws.Cell($"B{startrow}").Address, ws.Cell($"F{lrow2}").Address);

                range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                ws.Columns().AdjustToContents();
                #endregion

                workbook.SaveAs(fpath);
                success = true;
            }
            catch (Exception ex)
            {

            }

            return success;
        }

        public static bool WriteReportDetailsDERRIMON(string sheetname, string fpath, int startrow, DataGridView dataGrid, bool ismatched)
        {
            bool success = false;
            decimal totAmt = 0, totBal = 0;
            int lrow = 0, lrow2 = 0;
            XLWorkbook workbook = new XLWorkbook(fpath);

            IXLWorksheet ws = workbook.Worksheet(sheetname);

            try
            {
                #region WRITE DETAILS
                foreach (DataGridViewRow r in dataGrid.Rows)
                {
                    //Look for unmatched statements
                    if (Convert.ToBoolean(r.Cells["chbIsMatch"].Value) == ismatched)
                    {
                        //Write data to column B
                        //ws.Cell($"B{startrow}").Value = r.Cells["ReferenceNo"].Value.ToString();
                        ws.Cell($"B{startrow}").Value = r.Cells[dataGrid.Columns[1].HeaderText].Value.ToString();
                        //Write data to column C
                        ws.Cell($"C{startrow}").Value = r.Cells[dataGrid.Columns[2].HeaderText].Value.ToString();
                        //Write data to column D
                        ws.Cell($"D{startrow}").Value = r.Cells[dataGrid.Columns[3].HeaderText].Value.ToString();
                        //Write data to column E
                        ws.Cell($"E{startrow}").Value = r.Cells[dataGrid.Columns[4].HeaderText].Value.ToString().Replace(",", "");
                        ws.Cell($"E{startrow}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
                        totAmt = totAmt + Convert.ToDecimal(r.Cells[dataGrid.Columns[4].HeaderText].Value.ToString().Replace(",", ""));

                        startrow += 1;
                    }
                }
                #endregion

                #region STYLE REPORT
                //Styling
                lrow = ws.LastRowUsed().RowNumber();
                ws.Cell($"E{lrow + 1}").Value = totAmt.ToString();
                ws.Cell($"E{lrow + 1}").Style.Font.Bold = true;
                ws.Cell($"E{lrow + 1}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";


                int newRow = lrow + 1;
                ws.Range($"B{newRow}:E{newRow}").Style.Border.TopBorder = XLBorderStyleValues.Thin;

                lrow2 = ws.LastRowUsed().RowNumber();

                IXLRange range = ws.Range(ws.Cell($"B{startrow}").Address, ws.Cell($"E{lrow2}").Address);

                range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                ws.Columns().AdjustToContents();
                #endregion

                workbook.SaveAs(fpath);
                success = true;
            }
            catch (Exception ex)
            {

            }

            return success;
        }
        public static bool WriteReportDetailsV1(string sheetname, string fpath, int startrow, DataGridView dataGrid, bool ismatched)
        {
            bool success = false;
            decimal totAmt = 0, totBal = 0;
            int lrow = 0, lrow2 = 0;
            XLWorkbook workbook = new XLWorkbook(fpath);

            IXLWorksheet ws = workbook.Worksheet(sheetname);

            try
            {
                #region WRITE DETAILS
                foreach (DataGridViewRow r in dataGrid.Rows)
                {
                    //Look for unmatched statements
                    if (Convert.ToBoolean(r.Cells["chbIsMatch"].Value) == ismatched)
                    {
                        //Write data to column B
                        //ws.Cell($"B{startrow}").Value = r.Cells["ReferenceNo"].Value.ToString();
                        ws.Cell($"B{startrow}").Value = r.Cells[dataGrid.Columns[1].HeaderText].Value.ToString();
                        //Write data to column C
                        ws.Cell($"C{startrow}").Value = r.Cells[dataGrid.Columns[2].HeaderText].Value.ToString();
                        //Write data to column D
                        ws.Cell($"D{startrow}").Value = r.Cells[dataGrid.Columns[3].HeaderText].Value.ToString();
                        //Write data to column E
                        ws.Cell($"E{startrow}").Value = r.Cells[dataGrid.Columns[4].HeaderText].Value.ToString().Replace(",", "");
                        ws.Cell($"E{startrow}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
                        totAmt = totAmt + Convert.ToDecimal(r.Cells[dataGrid.Columns[4].HeaderText].Value.ToString().Replace(",", ""));

                        startrow += 1;
                    }
                }
                #endregion

                #region STYLE REPORT
                //Styling
                lrow = ws.LastRowUsed().RowNumber();
                ws.Cell($"E{lrow + 1}").Value = totAmt.ToString();
                ws.Cell($"E{lrow + 1}").Style.Font.Bold = true;
                ws.Cell($"E{lrow + 1}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";


                int newRow = lrow + 1;
                ws.Range($"B{newRow}:E{newRow}").Style.Border.TopBorder = XLBorderStyleValues.Thin;

                lrow2 = ws.LastRowUsed().RowNumber();

                IXLRange range = ws.Range(ws.Cell($"B{startrow}").Address, ws.Cell($"E{lrow2}").Address);

                range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                ws.Columns().AdjustToContents();
                #endregion

                workbook.SaveAs(fpath);
                success = true;
            }
            catch (Exception ex)
            {

            }

            return success;
        }

        public static bool WriteReportDetailsV2(string sheetname, string fpath, int startrow, DataGridView dataGrid, bool ismatched)
        {
            bool success = false;
            decimal totAmt = 0, totBal = 0;
            int lrow = 0, lrow2 = 0;
            XLWorkbook workbook = new XLWorkbook(fpath);

            IXLWorksheet ws = workbook.Worksheet(sheetname);

            try
            {
                #region WRITE DETAILS
                foreach (DataGridViewRow r in dataGrid.Rows)
                {
                    //Look for unmatched statements
                    if (Convert.ToBoolean(r.Cells["chbIsMatch"].Value) == ismatched)
                    {
                        //Write data to column B
                        //ws.Cell($"B{startrow}").Value = r.Cells["ReferenceNo"].Value.ToString();
                        ws.Cell($"B{startrow}").Value = r.Cells[dataGrid.Columns[1].HeaderText].Value.ToString();
                        //Write data to column C
                        ws.Cell($"C{startrow}").Value = r.Cells[dataGrid.Columns[2].HeaderText].Value.ToString();
                        //Write data to column D
                        ws.Cell($"D{startrow}").Value = r.Cells[dataGrid.Columns[3].HeaderText].Value.ToString();
                        //Write data to column E
                        ws.Cell($"E{startrow}").Value = r.Cells[dataGrid.Columns[4].HeaderText].Value.ToString().Replace(",", "");
                        ws.Cell($"E{startrow}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
                        totAmt = totAmt + Convert.ToDecimal(r.Cells[dataGrid.Columns[4].HeaderText].Value.ToString().Replace(",", ""));
                        //Write data to column F
                        ws.Cell($"F{startrow}").Value = r.Cells[dataGrid.Columns[5].HeaderText].Value.ToString().Replace(",", "");
                        ws.Cell($"F{startrow}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
                        totBal = totBal + Convert.ToDecimal(r.Cells[dataGrid.Columns[5].HeaderText].Value.ToString().Replace(",", ""));

                        startrow += 1;
                    }
                }
                #endregion

                #region STYLE REPORT
                //Styling
                lrow = ws.LastRowUsed().RowNumber();
                ws.Cell($"E{lrow + 1}").Value = totAmt.ToString();
                ws.Cell($"E{lrow + 1}").Style.Font.Bold = true;
                ws.Cell($"E{lrow + 1}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
                ws.Cell($"F{lrow + 1}").Value = totBal.ToString();
                ws.Cell($"F{lrow + 1}").Style.Font.Bold = true;
                ws.Cell($"F{lrow + 1}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";

                int newRow = lrow + 1;
                ws.Range($"B{newRow}:F{newRow}").Style.Border.TopBorder = XLBorderStyleValues.Thin;

                lrow2 = ws.LastRowUsed().RowNumber();

                IXLRange range = ws.Range(ws.Cell($"B{startrow}").Address, ws.Cell($"F{lrow2}").Address);

                range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                ws.Columns().AdjustToContents();
                #endregion

                workbook.SaveAs(fpath);
                success = true;
            }
            catch (Exception ex)
            {

            }

            return success;
        }

        public static bool WriteReportDetailsV3(string sheetname, string fpath, int startrow, DataGridView dataGrid, bool ismatched)
        {
            bool success = false;
            decimal totCRAmt = 0, totDBAmt = 0, totBal = 0;
            int lrow = 0, lrow2 = 0;
            XLWorkbook workbook = new XLWorkbook(fpath);

            IXLWorksheet ws = workbook.Worksheet(sheetname);

            try
            {
                #region WRITE DETAILS
                foreach (DataGridViewRow r in dataGrid.Rows)
                {
                    //Look for unmatched statements
                    if (Convert.ToBoolean(r.Cells["chbIsMatch"].Value) == ismatched)
                    {
                        //Write data to column B
                        //ws.Cell($"B{startrow}").Value = r.Cells["ReferenceNo"].Value.ToString();
                        ws.Cell($"B{startrow}").Value = r.Cells[dataGrid.Columns[1].HeaderText].Value.ToString();
                        //Write data to column C
                        ws.Cell($"C{startrow}").Value = r.Cells[dataGrid.Columns[2].HeaderText].Value.ToString();
                        //Write data to column D
                        ws.Cell($"D{startrow}").Value = r.Cells[dataGrid.Columns[3].HeaderText].Value.ToString();
                        //Write data to column E
                        ws.Cell($"E{startrow}").Value = r.Cells[dataGrid.Columns[4].HeaderText].Value.ToString().Replace(",", "");
                        ws.Cell($"E{startrow}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
                        totDBAmt = totDBAmt + Convert.ToDecimal(r.Cells[dataGrid.Columns[4].HeaderText].Value.ToString().Replace(",", ""));
                        //Write data to column F
                        ws.Cell($"F{startrow}").Value = r.Cells[dataGrid.Columns[5].HeaderText].Value.ToString().Replace(",", "");
                        ws.Cell($"F{startrow}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
                        totCRAmt = totCRAmt + Convert.ToDecimal(r.Cells[dataGrid.Columns[5].HeaderText].Value.ToString().Replace(",", ""));
                        //Write data to column F
                        ws.Cell($"G{startrow}").Value = r.Cells[dataGrid.Columns[6].HeaderText].Value.ToString().Replace(",", "");
                        ws.Cell($"G{startrow}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
                        totBal = totBal + Convert.ToDecimal(r.Cells[dataGrid.Columns[6].HeaderText].Value.ToString().Replace(",", ""));

                        startrow += 1;
                    }
                }
                #endregion

                #region STYLE REPORT
                //Styling
                lrow = ws.LastRowUsed().RowNumber();
                ws.Cell($"E{lrow + 1}").Value = totDBAmt.ToString();
                ws.Cell($"E{lrow + 1}").Style.Font.Bold = true;
                ws.Cell($"E{lrow + 1}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
                ws.Cell($"F{lrow + 1}").Value = totCRAmt.ToString();
                ws.Cell($"F{lrow + 1}").Style.Font.Bold = true;
                ws.Cell($"F{lrow + 1}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
                ws.Cell($"G{lrow + 1}").Value = totBal.ToString();
                ws.Cell($"G{lrow + 1}").Style.Font.Bold = true;
                ws.Cell($"G{lrow + 1}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";

                int newRow = lrow + 1;
                ws.Range($"B{newRow}:F{newRow}").Style.Border.TopBorder = XLBorderStyleValues.Thin;

                lrow2 = ws.LastRowUsed().RowNumber();

                IXLRange range = ws.Range(ws.Cell($"B{startrow}").Address, ws.Cell($"G{lrow2}").Address);

                range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                ws.Columns().AdjustToContents();
                #endregion

                workbook.SaveAs(fpath);
                success = true;
            }
            catch (Exception ex)
            {

            }

            return success;
        }

        //WITH LOCATION 
        public static bool WriteReportDetailsV4(string sheetname, string fpath, int startrow, DataGridView dataGrid, bool ismatched)
        {
            bool success = false;
            decimal totAmt = 0, totBal = 0;
            int lrow = 0, lrow2 = 0;
            XLWorkbook workbook = new XLWorkbook(fpath);

            IXLWorksheet ws = workbook.Worksheet(sheetname);

            try
            {
                #region WRITE DETAILS
                foreach (DataGridViewRow r in dataGrid.Rows)
                {
                    //Look for unmatched statements
                    if (Convert.ToBoolean(r.Cells["chbIsMatch"].Value) == ismatched)
                    {
                        //Write data to column B
                        //ws.Cell($"B{startrow}").Value = r.Cells["ReferenceNo"].Value.ToString();
                        ws.Cell($"B{startrow}").Value = r.Cells[dataGrid.Columns[1].HeaderText].Value.ToString();
                        //Write data to column C
                        ws.Cell($"C{startrow}").Value = r.Cells[dataGrid.Columns[2].HeaderText].Value.ToString();
                        //Write data to column D
                        ws.Cell($"D{startrow}").Value = r.Cells[dataGrid.Columns[3].HeaderText].Value.ToString();
                        //Write data to column E
                        ws.Cell($"E{startrow}").Value = r.Cells[dataGrid.Columns[4].HeaderText].Value.ToString().Replace(",", "");
                        ws.Cell($"E{startrow}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
                        totAmt = totAmt + Convert.ToDecimal(r.Cells[dataGrid.Columns[4].HeaderText].Value.ToString().Replace(",", ""));

                        ws.Cell($"F{startrow}").Value = r.Cells[dataGrid.Columns[5].HeaderText].Value.ToString().Replace(",", "");

                        startrow += 1;
                    }
                }
                #endregion

                #region STYLE REPORT
                //Styling
                lrow = ws.LastRowUsed().RowNumber();
                ws.Cell($"E{lrow + 1}").Value = totAmt.ToString();
                ws.Cell($"E{lrow + 1}").Style.Font.Bold = true;
                ws.Cell($"E{lrow + 1}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";


                int newRow = lrow + 1;
                ws.Range($"B{newRow}:F{newRow}").Style.Border.TopBorder = XLBorderStyleValues.Thin;

                lrow2 = ws.LastRowUsed().RowNumber();

                IXLRange range = ws.Range(ws.Cell($"B{startrow}").Address, ws.Cell($"F{lrow2}").Address);

                range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                ws.Columns().AdjustToContents();
                #endregion

                workbook.SaveAs(fpath);
                success = true;
            }
            catch (Exception ex)
            {

            }

            return success;
        }
    }

}
