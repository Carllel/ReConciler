using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using LinqToExcel;
using MetroSet_UI.Forms;
using ReConciler.Data;
using ReConciler.Model;
using ReConciler.Util;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = System.Windows.Forms.Application;

namespace ReConciler
{
    public partial class frmLoadStatement : MetroSetForm
    {
        public List<string> GetColumnHeaderNames(DataGridView dataGrid)
        {
            List<string> columnHeaderNames = new List<string>();
            int colindex = 1;
            try
            {
                foreach(var column in dataGrid.Columns)
                {
                   
                    var name = dataGrid.Columns[colindex].HeaderText;
                    columnHeaderNames.Add(name);
                    colindex++;

                }
            }
            catch(Exception ex)
            {

            }


            return columnHeaderNames;
        }
        public string vendor = "";

        #region VALIDATE VENDOR STATEMENT
        public async void ValidateStatement(List<SAPMaster> SAPMasterData, string vendor)
        {
            try
            {
                progressBar1.Visible = true;
                lblProgressPercent.Visible = true;
                (Application.OpenForms["Main"] as Main).slblStatus.Text = "Validation in progress...";
                var progress = new Progress<ProgressReport>();
                progress.ProgressChanged += (o, report) =>
                {
                    lblProgressPercent.Text = $"Processing...{report.PercentageComplete}%";
                    lblProgressPercent.Font = new Font(lblProgressPercent.Font, FontStyle.Bold);
                    lblProgressPercent.ForeColor = System.Drawing.Color.OrangeRed;

                    msBadgeTotalRec.BadgeText = $"{report.TotalRecords}";
                    msBadgeTotalMatch.BadgeText = $"{report.TotalMatched}";
                    msBadgeTotalNotMatch.BadgeText = $"{report.TotalNotMatched}";


                    progressBar1.Value = report.PercentageComplete;
                    progressBar1.Update();
                };

                switch (vendor)
                {
                    case "KIRK":
                        await ValidateSTMTEntriesKIRK(progress, SAPMasterData);
                        break;
                    case "DERRIMON":
                        await ValidateSTMTEntriesDERRIMON(progress, SAPMasterData);
                        break;
                    case "MASSY":
                        await ValidateSTMTEntriesMASSY(progress, SAPMasterData);
                        break;
                    case "CONSOLBAKERIES":
                        await ValidateSTMTEntriesConsolBakeries(progress, SAPMasterData);
                        break;
                    case "COPPERWOOD":
                        await ValidateSTMTEntriesCOPPERWOOD(progress, SAPMasterData);
                        break;
                    case "TGEDDES":
                        await ValidateSTMTEntriesTGEDDES(progress, SAPMasterData);
                        break;
                    case "FACEY":
                        await ValidateSTMTEntriesFACEY(progress, SAPMasterData);
                        break;
                    default:
                        break;
                }


                lblProgressPercent.Text = $"Done!";
                (Application.OpenForms["Main"] as Main).slblStatus.Text = "Validation Complete!";
                //Show Report Action
                if (msBadgeTotalNotMatch.BadgeText != "" || msBadgeTotalMatch.BadgeText != "")
                    btnGenerateReports.Visible = true;
                else
                    btnGenerateReports.Visible = false;
                MetroSetMessageBox.Show(this, "Validation Complete!.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MetroSetMessageBox.Show(this, $"{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region FACEY
        private Task ValidateSTMTEntriesFACEY(IProgress<ProgressReport> progress, List<SAPMaster> SAPMasterData)
        {
            int cntMatch = 0, cntNotMatch = 0, cnt = 1, totalProcess = dgSTSMEntries.Rows.Count;
            var progressReport = new ProgressReport();

            dgSTSMEntries.ClearSelection();

            var lstStmtEntries = dgSTSMEntries.DataSource as List<vendorStatement>;
            var SummlstStmtEntries = lstStmtEntries
                                    .GroupBy(l => l.ReferenceNo)
                                    .Select(cl => new validateVendorStatement
                                    {
                                        RefNo = cl.Last().ReferenceNo,
                                        DocDate = cl.Last().DocumentDate,
                                        Amt = cl.Sum(c => Convert.ToDecimal(c.Amount)).ToString(),
                                        Bal = cl.Sum(c => Convert.ToDecimal(c.Balance)).ToString()
                                    }).ToList();

            try
            {
                return Task.Run(() =>
                {
                    SummlstStmtEntries.ForEach(r => {
                        progressReport.PercentageComplete = cnt++ * 100 / SummlstStmtEntries.Count();

                        if (SAPMasterData.Any(s => s.Reference.Trim().StartsWith(r.RefNo.Trim()) &&
                                       s.Amount.ToString().Trim().Replace("-", "").Replace(",", "").StartsWith(r.Bal.Split('.')[0].Replace("-", "").Replace(",", ""))))
                        {
                            foreach (DataGridViewRow row in dgSTSMEntries.Rows)
                            {
                                if (row.Cells[1].Value.ToString().Equals(r.RefNo))
                                {
                                    row.Cells["chbIsMatch"].Value = CheckState.Checked;
                                    row.Cells["chbIsMatch"].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[1].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[2].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[3].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[4].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[5].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    cntMatch += 1;
                                }
                            }
                        }
                        else
                        {
                            foreach (DataGridViewRow row in dgSTSMEntries.Rows)
                            {
                                if (row.Cells[1].Value.ToString().Equals(r.RefNo))
                                {
                                    row.Cells["chbIsMatch"].Value = CheckState.Unchecked;
                                    row.Cells["chbIsMatch"].Style.BackColor = System.Drawing.Color.MistyRose;
                                    cntNotMatch += 1;
                                }
                            }
                        }

                        progressReport.TotalRecords = dgSTSMEntries.Rows.Count;
                        progressReport.TotalMatched = cntMatch;
                        progressReport.TotalNotMatched = cntNotMatch;

                        progress.Report(progressReport);
                        //Thread.Sleep(100);
                    });

                });
            }
            catch (ThreadInterruptedException e)
            {
                MetroSetMessageBox.Show(this, $"{e.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //MessageBox.Show(e.Message);
                return null;
            }
        }
        #endregion

        #region TGEDDES
        private Task ValidateSTMTEntriesTGEDDES(IProgress<ProgressReport> progress, List<SAPMaster> SAPMasterData)
        {
            int cntMatch = 0, cntNotMatch = 0, cnt = 1, totalProcess = dgSTSMEntries.Rows.Count;
            var progressReport = new ProgressReport();

            dgSTSMEntries.ClearSelection();

            var lstStmtEntries = dgSTSMEntries.DataSource as List<vendorStatement>;
            var SummlstStmtEntries = lstStmtEntries
                                    .GroupBy(l => l.ReferenceNo)
                                    .Select(cl => new validateVendorStatement
                                    {
                                        RefNo = cl.Last().ReferenceNo,
                                        DocDate = cl.Last().DocumentDate,
                                        Amt = cl.Sum(c => Convert.ToDecimal(c.Amount)).ToString(),
                                        Bal = cl.Sum(c => Convert.ToDecimal(c.Balance)).ToString()
                                    }).ToList();

            try
            {
                return Task.Run(() =>
                {
                    SummlstStmtEntries.ForEach(r => {
                        progressReport.PercentageComplete = cnt++ * 100 / SummlstStmtEntries.Count();

                        if (SAPMasterData.Any(s => s.Reference.Trim().StartsWith(r.RefNo.Trim()) &&
                                       s.Amount.ToString().Trim().Replace("-", "").Replace(",", "").StartsWith(r.Bal.Split('.')[0].Replace("-", "").Replace(",", ""))))
                        {
                            foreach (DataGridViewRow row in dgSTSMEntries.Rows)
                            {
                                if (row.Cells[1].Value.ToString().Equals(r.RefNo))
                                {
                                    row.Cells["chbIsMatch"].Value = CheckState.Checked;
                                    row.Cells["chbIsMatch"].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[1].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[2].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[3].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[4].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[5].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    cntMatch += 1;
                                }
                            }
                        }
                        else
                        {
                            foreach (DataGridViewRow row in dgSTSMEntries.Rows)
                            {
                                if (row.Cells[1].Value.ToString().Equals(r.RefNo))
                                {
                                    row.Cells["chbIsMatch"].Value = CheckState.Unchecked;
                                    row.Cells["chbIsMatch"].Style.BackColor = System.Drawing.Color.MistyRose;
                                    cntNotMatch += 1;
                                }
                            }
                        }
                        
                        progressReport.TotalRecords = dgSTSMEntries.Rows.Count;
                        progressReport.TotalMatched = cntMatch;
                        progressReport.TotalNotMatched = cntNotMatch;

                        progress.Report(progressReport);
                        Thread.Sleep(100);
                    });

                });
            }
            catch (ThreadInterruptedException e)
            {
                MetroSetMessageBox.Show(this, $"{e.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //MessageBox.Show(e.Message);
                return null;
            }
        }
        #endregion

        #region CONSOLIDATED BAKERIES
        private Task ValidateSTMTEntriesConsolBakeries(IProgress<ProgressReport> progress, List<SAPMaster> SAPMasterData)
        {
            int cntMatch = 0, cntNotMatch = 0, cnt = 1, totalProcess = dgSTSMEntries.Rows.Count;
            var progressReport = new ProgressReport();

            dgSTSMEntries.ClearSelection();

            var lstStmtEntries = dgSTSMEntries.DataSource as List<vendorSTMT>;
            var SummlstStmtEntries = lstStmtEntries
                                    .GroupBy(l => l.ReferenceNo)
                                    .Select(cl => new validateVendorSTMT
                                    {
                                        RefNo = cl.Last().ReferenceNo,
                                        DocDate = cl.Last().DocumentDate,
                                        Amt = cl.Sum(c => Convert.ToDecimal(c.Amount)).ToString(),
                                    }).ToList();

            try
            {
                return Task.Run(() =>
                {
                    SummlstStmtEntries.ForEach(r => {
                        progressReport.PercentageComplete = cnt++ * 100 / SummlstStmtEntries.Count();

                        if (SAPMasterData.Any(s => s.Reference.Trim().StartsWith(r.RefNo.Trim()) &&
                                       s.Amount.ToString().Trim().Replace("-", "").Replace(",", "").StartsWith(r.Amt.Split('.')[0].Replace("-", "").Replace(",", ""))))
                        {
                            foreach (DataGridViewRow row in dgSTSMEntries.Rows)
                            {
                                if (row.Cells[1].Value.ToString().Equals(r.RefNo))
                                {
                                    row.Cells["chbIsMatch"].Value = CheckState.Checked;
                                    row.Cells["chbIsMatch"].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[1].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[2].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[3].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[4].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    cntMatch += 1;
                                }
                            }
                        }
                        else
                        {
                            foreach (DataGridViewRow row in dgSTSMEntries.Rows)
                            {
                                if (row.Cells[1].Value.ToString().Equals(r.RefNo))
                                {
                                    row.Cells["chbIsMatch"].Value = CheckState.Unchecked;
                                    row.Cells["chbIsMatch"].Style.BackColor = System.Drawing.Color.MistyRose;
                                    cntNotMatch += 1;
                                }
                            }
                        }

                        progressReport.TotalRecords = dgSTSMEntries.Rows.Count;
                        progressReport.TotalMatched = cntMatch;
                        progressReport.TotalNotMatched = cntNotMatch;

                        progress.Report(progressReport);
                        //Thread.Sleep(100);
                    });

                });
            }
            catch (ThreadInterruptedException e)
            {
                MetroSetMessageBox.Show(this, $"{e.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //MessageBox.Show(e.Message);
                return null;
            }
        }

        #endregion

        #region MASSY

        private Task ValidateSTMTEntriesMASSY(IProgress<ProgressReport> progress, List<SAPMaster> SAPMasterData)
        {
            int cntMatch = 0, cntNotMatch = 0, cnt = 1, totalProcess = dgSTSMEntries.Rows.Count;
            var progressReport = new ProgressReport();

            dgSTSMEntries.ClearSelection();

            var lstStmtEntries = dgSTSMEntries.DataSource as List<vendorSTMT>;
            var SummlstStmtEntries = lstStmtEntries
                                    .GroupBy(l => l.ReferenceNo)
                                    .Select(cl => new validateVendorSTMT
                                    {
                                        RefNo = cl.Last().ReferenceNo,
                                        DocDate = cl.Last().DocumentDate,
                                        Amt = cl.Sum(c => Convert.ToDecimal(c.Amount)).ToString(),
                                    }).ToList();

            try
            {
                return Task.Run(() =>
                {
                    SummlstStmtEntries.ForEach(r => {
                        progressReport.PercentageComplete = cnt++ * 100 / SummlstStmtEntries.Count();

                        if (SAPMasterData.Any(s => s.Reference.Trim().StartsWith(r.RefNo.Trim()) &&
                                       s.Amount.ToString().Trim().Replace("-", "").Replace(",", "").StartsWith(r.Amt.Split('.')[0].Replace("-", "").Replace(",", ""))))
                        {
                            foreach (DataGridViewRow row in dgSTSMEntries.Rows)
                            {
                                if (row.Cells[1].Value.ToString().Equals(r.RefNo))
                                {
                                    row.Cells["chbIsMatch"].Value = CheckState.Checked;
                                    row.Cells["chbIsMatch"].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[1].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[2].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[3].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[4].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    cntMatch += 1;
                                }
                            }
                        }
                        else
                        {
                            foreach (DataGridViewRow row in dgSTSMEntries.Rows)
                            {
                                if (row.Cells[1].Value.ToString().Equals(r.RefNo))
                                {
                                    row.Cells["chbIsMatch"].Value = CheckState.Unchecked;
                                    row.Cells["chbIsMatch"].Style.BackColor = System.Drawing.Color.MistyRose;
                                    cntNotMatch += 1;
                                }
                            }
                        }

                        progressReport.TotalRecords = dgSTSMEntries.Rows.Count;
                        progressReport.TotalMatched = cntMatch;
                        progressReport.TotalNotMatched = cntNotMatch;

                        progress.Report(progressReport);
                        //Thread.Sleep(100);
                    });

                });
            }
            catch (ThreadInterruptedException e)
            {
                MetroSetMessageBox.Show(this, $"{e.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //MessageBox.Show(e.Message);
                return null;
            }
        }
        
        #endregion

        #region DERRIMON

        private Task ValidateSTMTEntriesDERRIMON(IProgress<ProgressReport> progress, List<SAPMaster> SAPMasterData)
        {
            int cntMatch = 0, cntNotMatch = 0, cnt = 1, totalProcess = dgSTSMEntries.Rows.Count;
            var progressReport = new ProgressReport();

            dgSTSMEntries.ClearSelection();

            var lstStmtEntries = dgSTSMEntries.DataSource as List<vendorSTMT>;
            var SummlstStmtEntries = lstStmtEntries
                                    .GroupBy(l => l.ReferenceNo)
                                    .Select(cl => new validateVendorSTMT
                                    {
                                        RefNo = cl.Last().ReferenceNo,
                                        DocDate = cl.Last().DocumentDate,
                                        Amt = cl.Sum(c => Convert.ToDecimal(c.Amount)).ToString(),
                                    }).ToList();

            try
            {
                return Task.Run(() =>
                {
                    SummlstStmtEntries.ForEach(r => {
                        progressReport.PercentageComplete = cnt++ * 100 / SummlstStmtEntries.Count();

                        if (SAPMasterData.Any(s => s.Reference.Trim().Contains(r.RefNo.Trim().Substring(r.RefNo.Trim().Length - 5)) &&
                                       s.Amount.ToString().Trim().Replace("-", "").Replace(",", "").StartsWith(r.Amt.Split('.')[0].Replace("-", "").Replace(",", ""))))
                        {
                            foreach (DataGridViewRow row in dgSTSMEntries.Rows)
                            {
                                if (row.Cells[1].Value.ToString().Equals(r.RefNo))
                                {
                                    row.Cells["chbIsMatch"].Value = CheckState.Checked;
                                    row.Cells["chbIsMatch"].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[1].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[2].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[3].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[4].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    //row.Cells[dgSTSMEntries.Columns[5].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    cntMatch += 1;
                                }
                            }
                        }
                        else
                        {
                            foreach (DataGridViewRow row in dgSTSMEntries.Rows)
                            {
                                if (row.Cells[1].Value.ToString().Equals(r.RefNo))
                                {
                                    row.Cells["chbIsMatch"].Value = CheckState.Unchecked;
                                    row.Cells["chbIsMatch"].Style.BackColor = System.Drawing.Color.MistyRose;
                                    cntNotMatch += 1;
                                }
                            }
                        }

                        progressReport.TotalRecords = dgSTSMEntries.Rows.Count;
                        progressReport.TotalMatched = cntMatch;
                        progressReport.TotalNotMatched = cntNotMatch;

                        progress.Report(progressReport);
                        //Thread.Sleep(100);
                    });

                });
            }
            catch (ThreadInterruptedException e)
            {
                MetroSetMessageBox.Show(this, $"{e.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //MessageBox.Show(e.Message);
                return null;
            }
        }

        #endregion

        #region KIRK

        private Task ValidateSTMTEntriesKIRK(IProgress<ProgressReport> progress, List<SAPMaster> SAPMasterData)
        {
            int cntMatch = 0, cntNotMatch = 0, cnt = 1, totalProcess = dgSTSMEntries.Rows.Count;
            var progressReport = new ProgressReport();

            dgSTSMEntries.ClearSelection();

            var lstStmtEntries = dgSTSMEntries.DataSource as List<vendorSTMT>;
            var SummlstStmtEntries = lstStmtEntries
                                    .GroupBy(l => l.ReferenceNo)
                                    .Select(cl => new validateVendorSTMT
                                    {
                                        RefNo = cl.Last().ReferenceNo,
                                        DocDate = cl.Last().DocumentDate,
                                        Amt = cl.Sum(c => Convert.ToDecimal(c.Amount)).ToString(),
                                    }).ToList();

            try
            {
                return Task.Run(() =>
                {
                    SummlstStmtEntries.ForEach(r => {
                        progressReport.PercentageComplete = cnt++ * 100 / SummlstStmtEntries.Count();

                        if (SAPMasterData.Any(s => s.Reference.Trim().StartsWith(r.RefNo.Trim()) &&
                                       s.Amount.ToString().Trim().Replace("-", "").Replace(",", "").StartsWith(r.Amt.Split('.')[0].Replace("-", "").Replace(",", ""))))
                        {
                            foreach (DataGridViewRow row in dgSTSMEntries.Rows)
                            {
                                if (row.Cells[1].Value.ToString().Equals(r.RefNo))
                                {
                                    row.Cells["chbIsMatch"].Value = CheckState.Checked;
                                    row.Cells["chbIsMatch"].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[1].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[2].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[3].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[4].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;

                                    cntMatch += 1;
                                }
                            }
                        }
                        else
                        {
                            foreach (DataGridViewRow row in dgSTSMEntries.Rows)
                            {
                                if (row.Cells[1].Value.ToString().Equals(r.RefNo))
                                {
                                    row.Cells["chbIsMatch"].Value = CheckState.Unchecked;
                                    row.Cells["chbIsMatch"].Style.BackColor = System.Drawing.Color.MistyRose;
                                    cntNotMatch += 1;
                                }
                            }
                        }

                        progressReport.TotalRecords = dgSTSMEntries.Rows.Count;
                        progressReport.TotalMatched = cntMatch;
                        progressReport.TotalNotMatched = cntNotMatch;

                        progress.Report(progressReport);
                        //Thread.Sleep(100);
                    });

                });
            }
            catch (ThreadInterruptedException e)
            {
                MetroSetMessageBox.Show(this, $"{e.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //MessageBox.Show(e.Message);
                return null;
            }
        }

        #endregion

        #region COPPERWOOD
        private Task ValidateSTMTEntriesCOPPERWOOD(IProgress<ProgressReport> progress, List<SAPMaster> SAPMasterData)
        {
            int cntMatch = 0, cntNotMatch = 0, cnt = 1, totalProcess = dgSTSMEntries.Rows.Count;
            var progressReport = new ProgressReport();

            dgSTSMEntries.ClearSelection();

            var lstStmtEntries = dgSTSMEntries.DataSource as List<vendorStatement>;
            var SummlstStmtEntries = lstStmtEntries
                                    .GroupBy(l => l.ReferenceNo)
                                    .Select(cl => new validateVendorStatement
                                    {
                                        RefNo = cl.Last().ReferenceNo,
                                        DocDate = cl.Last().DocumentDate,
                                        Amt = cl.Sum(c => Convert.ToDecimal(c.Amount)).ToString(),
                                        Bal = cl.Sum(c => Convert.ToDecimal(c.Balance)).ToString()
                                    }).ToList();

            try
            {
                return Task.Run(() =>
                {
                    SummlstStmtEntries.ForEach(r => {
                        progressReport.PercentageComplete = cnt++ * 100 / SummlstStmtEntries.Count();

                        if (SAPMasterData.Any(s => s.Reference.Trim().StartsWith(r.RefNo.Trim()) && 
                                       s.Amount.ToString().Trim().Replace("-", "").Replace(",", "").StartsWith(r.Bal.Split('.')[0].Replace("-", "").Replace(",", "")) ))
                        {
                            foreach (DataGridViewRow row in dgSTSMEntries.Rows)
                            {
                                if (row.Cells[1].Value.ToString().Equals(r.RefNo))
                                {
                                    row.Cells["chbIsMatch"].Value = CheckState.Checked;
                                    row.Cells["chbIsMatch"].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[1].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[2].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[3].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[4].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;
                                    row.Cells[dgSTSMEntries.Columns[5].HeaderText].Style.BackColor = System.Drawing.Color.LightGreen;

                                    cntMatch += 1;
                                }
                            }
                        }
                        else
                        {
                            foreach (DataGridViewRow row in dgSTSMEntries.Rows)
                            {
                                if (row.Cells[1].Value.ToString().Equals(r.RefNo))
                                {
                                    row.Cells["chbIsMatch"].Value = CheckState.Unchecked;
                                    row.Cells["chbIsMatch"].Style.BackColor = System.Drawing.Color.MistyRose;
                                    cntNotMatch += 1;
                                }
                            }
                        }

                        progressReport.TotalRecords = dgSTSMEntries.Rows.Count;
                        progressReport.TotalMatched = cntMatch;
                        progressReport.TotalNotMatched = cntNotMatch;

                        progress.Report(progressReport);
                        //Thread.Sleep(100);
                    });

                });
            }
            catch (ThreadInterruptedException e)
            {
                MetroSetMessageBox.Show(this, $"{e.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //MessageBox.Show(e.Message);
                return null;
            }
        }
        //private Task ValidateSTMTEntriesCOPPERWOOD_OLD(IProgress<ProgressReport> progress, List<SAPMaster> SAPMasterData)
        //{
        //    int cntMatch = 0, cntNotMatch = 0, cnt = 1, totalProcess = dgSTSMEntries.Rows.Count;
        //    var progressReport = new ProgressReport();

        //    dgSTSMEntries.ClearSelection();

        //    try
        //    {
        //        return Task.Run(() =>
        //        {
        //            foreach (DataGridViewRow row in dgSTSMEntries.Rows)
        //            {
        //                progressReport.PercentageComplete = cnt++ * 100 / totalProcess;

        //                string docNo = row.Cells["ReferenceNo"].Value.ToString();
        //                decimal amount = Convert.ToDecimal(row.Cells["Balance"].Value.ToString());
        //                DateTime docDate = Convert.ToDateTime(row.Cells["DocumentDate"].Value.ToString());

        //                if (SAPMasterData.Any(s => s.Reference == docNo && s.Amount == amount && s.DocumentDate.Value.Date == docDate.Date))
        //                {
        //                    row.Cells["chbIsMatch"].Value = CheckState.Checked;
        //                    row.Cells["chbIsMatch"].Style.BackColor = System.Drawing.Color.LightGreen;
        //                    cntMatch += 1;
        //                }
        //                else
        //                {
        //                    row.Cells["chbIsMatch"].Value = CheckState.Unchecked;
        //                    row.Cells["chbIsMatch"].Style.BackColor = Color.MistyRose;
        //                    cntNotMatch += 1;
        //                }

        //                progressReport.TotalRecords = dgSTSMEntries.Rows.Count;
        //                progressReport.TotalMatched = cntMatch;
        //                progressReport.TotalNotMatched = cntNotMatch;

        //                progress.Report(progressReport);
        //                //Thread.Sleep(1000);
        //            }



        //        });
        //    }
        //    catch (ThreadInterruptedException e)
        //    {
        //        MetroSetMessageBox.Show(this, $"{e.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        //MessageBox.Show(e.Message);
        //        return null;
        //    }
        //}
        #endregion

        #region WISYNCO
        private Task ValidateSTMTEntriesWISYNCO(IProgress<ProgressReport> progress, List<SAPMaster> lstSAPMast)
        {
            int cntMatch = 0, cntNotMatch = 0, cnt = 1, totalProcess = dgSTSMEntries.Rows.Count;
            var progressReport = new ProgressReport();

            dgSTSMEntries.ClearSelection();

            try
            {
                return Task.Run(() =>
                {
                    foreach (DataGridViewRow row in dgSTSMEntries.Rows)
                    {
                        progressReport.PercentageComplete = cnt++ * 100 / totalProcess;
                        string referenceNo = row.Cells["ReferenceNo"].Value.ToString();
                        string drAmt = row.Cells["Debit"].Value.ToString();
                        string crAmt = row.Cells["Credit"].Value.ToString();

                        drAmt = drAmt.Length == 0 ? "0.00" : drAmt.Replace(",", "");//handle null fields
                        crAmt = crAmt.Length == 0 ? "0.00" : crAmt.Replace(",", "");//handle null fields
                        //You could use a nested Any() for this check which is available on any Enumerable:
                        //bool hasMatch = myStrings.Any(x => parameters.Any(y => y.source == x));

                        //if (lstSAPMast.Any(s=>s.DocumentNo == referenceNo))
                        //{
                        //    row.Cells["chbIsMatch"].Value = CheckState.Checked;
                        //    row.Cells["chbIsMatch"].Style.BackColor = System.Drawing.Color.LightGreen;
                        //    cntMatch += 1;
                        //}
                        //else
                        //{
                        //    row.Cells["chbIsMatch"].Value = CheckState.Unchecked;
                        //    row.Cells["chbIsMatch"].Style.BackColor = Color.MistyRose;
                        //    cntNotMatch += 1;
                        //}



                        ////Remove minus sign from the end of the amount and place it at hte start
                        //if (crAmt.Contains("-"))
                        //    crAmt = crAmt.Substring(crAmt.Length - 1) + crAmt.Remove(crAmt.Length - 1);

                        //if (Convert.ToDecimal(drAmt) >= 50000)
                        //{
                        //    row.Cells["chbIsMatch"].Value = CheckState.Checked;
                        //    row.Cells["chbIsMatch"].Style.BackColor = System.Drawing.Color.LightGreen;
                        //    cntMatch += 1;
                        //}
                        //else
                        //{
                        //    row.Cells["chbIsMatch"].Value = CheckState.Unchecked;
                        //    row.Cells["chbIsMatch"].Style.BackColor = Color.MistyRose;
                        //    cntNotMatch += 1;
                        //}

                        progressReport.TotalRecords = dgSTSMEntries.Rows.Count;
                        progressReport.TotalMatched = cntMatch;
                        progressReport.TotalNotMatched = cntNotMatch;

                        progress.Report(progressReport);
                        Thread.Sleep(100);
                    }



                });
            }
            catch (ThreadInterruptedException e)
            {

                MetroSetMessageBox.Show(this, $"{e.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        public async void ValidateWISYNCO(List<SAPMaster> SAPMasterData)
        {
            try
            {
                progressBar1.Visible = true;
                lblProgressPercent.Visible = true;
                (Application.OpenForms["Main"] as Main).slblStatus.Text = "Validation in progress...";

                var progress = new Progress<ProgressReport>();
                progress.ProgressChanged += (o, report) =>
                {
                    lblProgressPercent.Text = $"Processing...{report.PercentageComplete}%";
                    lblProgressPercent.Font = new Font(lblProgressPercent.Font, FontStyle.Bold);
                    lblProgressPercent.ForeColor = System.Drawing.Color.OrangeRed;

                    msBadgeTotalRec.BadgeText = $"{report.TotalRecords}";
                    //lblTotRecords.Font = new Font(lblTotRecords.Font, FontStyle.Bold);
                    //lblTotRecords.ForeColor = System.Drawing.Color.Blue;

                    msBadgeTotalMatch.BadgeText = $"{report.TotalMatched}";
                    //lblTotMach.Font = new Font(lblTotMach.Font, FontStyle.Bold);
                    //lblTotMach.ForeColor = System.Drawing.Color.Green;

                    msBadgeTotalNotMatch.BadgeText = $"{report.TotalNotMatched}";
                    //lblTotNotMatch.Font = new Font(lblTotNotMatch.Font, FontStyle.Bold);
                    //lblTotNotMatch.ForeColor = System.Drawing.Color.DarkRed;

                    progressBar1.Value = report.PercentageComplete;
                    progressBar1.Update();
                };

                await ValidateSTMTEntriesWISYNCO(progress,SAPMasterData);
                lblProgressPercent.Text = $"Done!";
                (Application.OpenForms["Main"] as Main).slblStatus.Text = "Validation Complete!";

                //Show Report Action
                if (msBadgeTotalNotMatch.BadgeText != "0" || msBadgeTotalNotMatch.BadgeText != "")
                    btnGenerateReports.Visible = true;
                else
                    btnGenerateReports.Visible = false;
                MetroSetMessageBox.Show(this, "Validation Complete!.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //MessageBox.Show("Validation Complete!.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion

        #endregion

        #region GENERATE REPORTS
        public void GenerateReportFACEY(string vendor)
        {
            //List<string> headerNames = new List<string> { "Transaction Reference", "Document Date", "Document Type", "Amount", "Balance" };
            var headerNames = GetColumnHeaderNames(dgSTSMEntries);

            try
            {

                string reportsDIR = Path.GetDirectoryName(txtStmtFileLoc.Text.Trim()) + $@"\REPORTS\{vendor}";
                //Create directoy if doesn't exist
                if (!Directory.Exists(reportsDIR))
                    Directory.CreateDirectory(reportsDIR);


                //Define Sheet Names
                string shtnameNOTMATCHED = "NOTMATCHED";
                string shtnameMATCHED = "MATCHED";

                string reportFileNme = reportsDIR + $@"\RPT_VAL_RESULTS_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";

                //create workbook
                var isCreated = UtilManager.CreateReportsWorkbook(reportFileNme, shtnameNOTMATCHED, shtnameMATCHED);

                if (isCreated)
                {
                    #region WRITE REPORTS
                    //NOT MATCHED
                    if (msBadgeTotalNotMatch.BadgeText != "0")
                    {
                        ////Report Header
                        var ishdrCreated = UtilManager.WriteReportHeaderData(shtnameNOTMATCHED, reportFileNme, "FACEY COMMODITY CO. LTD.", "B", "2", headerNames, 3);
                        if (ishdrCreated)
                        {
                            //Write details
                            #region WRITE DETAILS
                            var isDetailsWritten = UtilManager.WriteReportDetailsV2(shtnameNOTMATCHED, reportFileNme, 4, dgSTSMEntries, false);
                            #endregion
                        }
                    }
                    //MATCHED
                    if (msBadgeTotalMatch.BadgeText != "0")
                    {
                        ////Report Header
                        var ishdrCreated = UtilManager.WriteReportHeaderData(shtnameMATCHED, reportFileNme, "FACEY COMMODITY CO. LTD.", "B", "2", headerNames, 3);
                        if (ishdrCreated)
                        {
                            //Write details
                            #region WRITE DETAILS
                            var isDetailsWritten = UtilManager.WriteReportDetailsV2(shtnameMATCHED, reportFileNme, 4, dgSTSMEntries, true);
                            #endregion
                        }
                    }
                    #endregion

                    MetroSetMessageBox.Show(this, "Reports Generated Successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MetroSetMessageBox.Show(this, "Unable to generate report.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MetroSetMessageBox.Show(this, $"{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void GenerateReportTGEDDES(string vendor)
        {          
            //List<string> headerNames = new List<string> { "Transaction Reference", "Document Date", "Document Type", "Amount", "Balance" };
            var headerNames = GetColumnHeaderNames(dgSTSMEntries);

            try
            {

                string reportsDIR = Path.GetDirectoryName(txtStmtFileLoc.Text.Trim()) + $@"\REPORTS\{vendor}";
                //Create directoy if doesn't exist
                if (!Directory.Exists(reportsDIR))
                    Directory.CreateDirectory(reportsDIR);


                //Define Sheet Names
                string shtnameNOTMATCHED = "NOTMATCHED";
                string shtnameMATCHED = "MATCHED";

                string reportFileNme = reportsDIR + $@"\RPT_VAL_RESULTS_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";            

                //create workbook
                var isCreated = UtilManager.CreateReportsWorkbook(reportFileNme, shtnameNOTMATCHED,shtnameMATCHED);
                
                if (isCreated)
                {
                    #region WRITE REPORTS
                    //NOT MATCHED
                    if (msBadgeTotalNotMatch.BadgeText != "0")
                    {
                        ////Report Header
                        var ishdrCreated = UtilManager.WriteReportHeaderData(shtnameNOTMATCHED, reportFileNme, "T. GEDDES GRANT DIST. LTD.", "B", "2", headerNames, 3);
                        if (ishdrCreated)
                        {
                            //Write details
                            #region WRITE DETAILS
                            var isDetailsWritten = UtilManager.WriteReportDetailsTGEDDES(shtnameNOTMATCHED, reportFileNme, 4, dgSTSMEntries,false);
                            #endregion
                        }
                    }
                    //MATCHED
                    if (msBadgeTotalMatch.BadgeText != "0")
                    {
                        ////Report Header
                        var ishdrCreated = UtilManager.WriteReportHeaderData(shtnameMATCHED, reportFileNme, "T. GEDDES GRANT DIST. LTD.", "B", "2", headerNames, 3);
                        if (ishdrCreated)
                        {
                            //Write details
                            #region WRITE DETAILS
                            var isDetailsWritten = UtilManager.WriteReportDetailsTGEDDES(shtnameMATCHED, reportFileNme, 4, dgSTSMEntries,true);
                            #endregion
                        }
                    }
                    #endregion
                    
                    MetroSetMessageBox.Show(this, "Reports Generated Successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MetroSetMessageBox.Show(this, "Unable to generate report.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MetroSetMessageBox.Show(this, $"{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void GenerateReportCOPPERWOOD(string vendor)
        {
            var headerNames = GetColumnHeaderNames(dgSTSMEntries);

            try
            {

                string reportsDIR = Path.GetDirectoryName(txtStmtFileLoc.Text.Trim()) + $@"\REPORTS\{vendor}";
                //Create directoy if doesn't exist
                if (!Directory.Exists(reportsDIR))
                    Directory.CreateDirectory(reportsDIR);


                //Define Sheet Names
                string shtnameNOTMATCHED = "NOTMATCHED";
                string shtnameMATCHED = "MATCHED";

                string reportFileNme = reportsDIR + $@"\RPT_VAL_RESULTS_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";

                //create workbook
                var isCreated = UtilManager.CreateReportsWorkbook(reportFileNme, shtnameNOTMATCHED, shtnameMATCHED);

                if (isCreated)
                {
                    #region WRITE REPORTS
                    //NOT MATCHED
                    if (msBadgeTotalNotMatch.BadgeText != "0")
                    {
                        ////Report Header
                        var ishdrCreated = UtilManager.WriteReportHeaderData(shtnameNOTMATCHED, reportFileNme, "COPPERWOOD", "B", "2", headerNames, 3);
                        if (ishdrCreated)
                        {
                            //Write details
                            #region WRITE DETAILS
                            var isDetailsWritten = UtilManager.WriteReportDetailsCOPPERWOOD(shtnameNOTMATCHED, reportFileNme, 4, dgSTSMEntries, false);
                            #endregion
                        }
                    }
                    //MATCHED
                    if (msBadgeTotalMatch.BadgeText != "0")
                    {
                        ////Report Header
                        var ishdrCreated = UtilManager.WriteReportHeaderData(shtnameMATCHED, reportFileNme, "COPPERWOOD", "B", "2", headerNames, 3);
                        if (ishdrCreated)
                        {
                            //Write details
                            #region WRITE DETAILS
                            var isDetailsWritten = UtilManager.WriteReportDetailsCOPPERWOOD(shtnameMATCHED, reportFileNme, 4, dgSTSMEntries, true);
                            #endregion
                        }
                    }
                    #endregion

                    MetroSetMessageBox.Show(this, "Reports Generated Successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MetroSetMessageBox.Show(this, "Unable to generate report.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MetroSetMessageBox.Show(this, $"{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void GenerateReportDERRIMON(string vendor)
        {
            var headerNames = GetColumnHeaderNames(dgSTSMEntries);

            try
            {

                string reportsDIR = Path.GetDirectoryName(txtStmtFileLoc.Text.Trim()) + $@"\REPORTS\{vendor}";
                //Create directoy if doesn't exist
                if (!Directory.Exists(reportsDIR))
                    Directory.CreateDirectory(reportsDIR);


                //Define Sheet Names
                string shtnameNOTMATCHED = "NOTMATCHED";
                string shtnameMATCHED = "MATCHED";

                string reportFileNme = reportsDIR + $@"\RPT_VAL_RESULTS_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";

                //create workbook
                var isCreated = UtilManager.CreateReportsWorkbook(reportFileNme, shtnameNOTMATCHED, shtnameMATCHED);

                if (isCreated)
                {
                    #region WRITE REPORTS
                    //NOT MATCHED
                    if (msBadgeTotalNotMatch.BadgeText != "0")
                    {
                        ////Report Header
                        var ishdrCreated = UtilManager.WriteReportHeaderData(shtnameNOTMATCHED, reportFileNme, "DERRIMON TRADING LTD", "B", "2", headerNames, 3);
                        if (ishdrCreated)
                        {
                            //Write details
                            #region WRITE DETAILS
                            var isDetailsWritten = UtilManager.WriteReportDetailsV1(shtnameNOTMATCHED, reportFileNme, 4, dgSTSMEntries, false);
                            #endregion
                        }
                    }
                    //MATCHED
                    if (msBadgeTotalMatch.BadgeText != "0")
                    {
                        ////Report Header
                        var ishdrCreated = UtilManager.WriteReportHeaderData(shtnameMATCHED, reportFileNme, "DERRIMON TRADING LTD", "B", "2", headerNames, 3);
                        if (ishdrCreated)
                        {
                            //Write details
                            #region WRITE DETAILS
                            var isDetailsWritten = UtilManager.WriteReportDetailsV1(shtnameMATCHED, reportFileNme, 4, dgSTSMEntries, true);
                            #endregion
                        }
                    }
                    #endregion

                    MetroSetMessageBox.Show(this, "Reports Generated Successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MetroSetMessageBox.Show(this, "Unable to generate report.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MetroSetMessageBox.Show(this, $"{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void GenerateReportCONSOLBAKERIES(string vendor)
        {
            var headerNames = GetColumnHeaderNames(dgSTSMEntries);

            try
            {

                string reportsDIR = Path.GetDirectoryName(txtStmtFileLoc.Text.Trim()) + $@"\REPORTS\{vendor}";
                //Create directoy if doesn't exist
                if (!Directory.Exists(reportsDIR))
                    Directory.CreateDirectory(reportsDIR);


                //Define Sheet Names
                string shtnameNOTMATCHED = "NOTMATCHED";
                string shtnameMATCHED = "MATCHED";

                string reportFileNme = reportsDIR + $@"\RPT_VAL_RESULTS_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";

                //create workbook
                var isCreated = UtilManager.CreateReportsWorkbook(reportFileNme, shtnameNOTMATCHED, shtnameMATCHED);

                if (isCreated)
                {
                    #region WRITE REPORTS
                    //NOT MATCHED
                    if (msBadgeTotalNotMatch.BadgeText != "0")
                    {
                        ////Report Header
                        var ishdrCreated = UtilManager.WriteReportHeaderData(shtnameNOTMATCHED, reportFileNme, "CONSOLIDATED BAKERIES (JA) LTD", "B", "2", headerNames, 3);
                        if (ishdrCreated)
                        {
                            //Write details
                            #region WRITE DETAILS
                            var isDetailsWritten = UtilManager.WriteReportDetailsV1(shtnameNOTMATCHED, reportFileNme, 4, dgSTSMEntries, false);
                            #endregion
                        }
                    }
                    //MATCHED
                    if (msBadgeTotalMatch.BadgeText != "0")
                    {
                        ////Report Header
                        var ishdrCreated = UtilManager.WriteReportHeaderData(shtnameMATCHED, reportFileNme, "CONSOLIDATED BAKERIES (JA) LTD", "B", "2", headerNames, 3);
                        if (ishdrCreated)
                        {
                            //Write details
                            #region WRITE DETAILS
                            var isDetailsWritten = UtilManager.WriteReportDetailsV1(shtnameMATCHED, reportFileNme, 4, dgSTSMEntries, true);
                            #endregion
                        }
                    }
                    #endregion

                    MetroSetMessageBox.Show(this, "Reports Generated Successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MetroSetMessageBox.Show(this, "Unable to generate report.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MetroSetMessageBox.Show(this, $"{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void GenerateReportMASSY(string vendor)
        {
            var headerNames = GetColumnHeaderNames(dgSTSMEntries);

            try
            {

                string reportsDIR = Path.GetDirectoryName(txtStmtFileLoc.Text.Trim()) + $@"\REPORTS\{vendor}";
                //Create directoy if doesn't exist
                if (!Directory.Exists(reportsDIR))
                    Directory.CreateDirectory(reportsDIR);


                //Define Sheet Names
                string shtnameNOTMATCHED = "NOTMATCHED";
                string shtnameMATCHED = "MATCHED";

                string reportFileNme = reportsDIR + $@"\RPT_VAL_RESULTS_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";

                //create workbook
                var isCreated = UtilManager.CreateReportsWorkbook(reportFileNme, shtnameNOTMATCHED, shtnameMATCHED);

                if (isCreated)
                {
                    #region WRITE REPORTS
                    //NOT MATCHED
                    if (msBadgeTotalNotMatch.BadgeText != "0")
                    {
                        ////Report Header
                        var ishdrCreated = UtilManager.WriteReportHeaderData(shtnameNOTMATCHED, reportFileNme, "MASSY DISTRIBUTION", "B", "2", headerNames, 3);
                        if (ishdrCreated)
                        {
                            //Write details
                            #region WRITE DETAILS
                            var isDetailsWritten = UtilManager.WriteReportDetailsV1(shtnameNOTMATCHED, reportFileNme, 4, dgSTSMEntries, false);
                            #endregion
                        }
                    }
                    //MATCHED
                    if (msBadgeTotalMatch.BadgeText != "0")
                    {
                        ////Report Header
                        var ishdrCreated = UtilManager.WriteReportHeaderData(shtnameMATCHED, reportFileNme, "MASSY DISTRIBUTION", "B", "2", headerNames, 3);
                        if (ishdrCreated)
                        {
                            //Write details
                            #region WRITE DETAILS
                            var isDetailsWritten = UtilManager.WriteReportDetailsV1(shtnameMATCHED, reportFileNme, 4, dgSTSMEntries, true);
                            #endregion
                        }
                    }
                    #endregion

                    MetroSetMessageBox.Show(this, "Reports Generated Successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MetroSetMessageBox.Show(this, "Unable to generate report.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MetroSetMessageBox.Show(this, $"{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void GenerateReportKIRK(string vendor)
        {
            var headerNames = GetColumnHeaderNames(dgSTSMEntries);

            try
            {

                string reportsDIR = Path.GetDirectoryName(txtStmtFileLoc.Text.Trim()) + $@"\REPORTS\{vendor}";
                //Create directoy if doesn't exist
                if (!Directory.Exists(reportsDIR))
                    Directory.CreateDirectory(reportsDIR);


                //Define Sheet Names
                string shtnameNOTMATCHED = "NOTMATCHED";
                string shtnameMATCHED = "MATCHED";

                string reportFileNme = reportsDIR + $@"\RPT_VAL_RESULTS_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";

                //create workbook
                var isCreated = UtilManager.CreateReportsWorkbook(reportFileNme, shtnameNOTMATCHED, shtnameMATCHED);

                if (isCreated)
                {
                    #region WRITE REPORTS
                    //NOT MATCHED
                    if (msBadgeTotalNotMatch.BadgeText != "0")
                    {
                        ////Report Header
                        var ishdrCreated = UtilManager.WriteReportHeaderData(shtnameNOTMATCHED, reportFileNme, "KIRK DISTRIBUTORS", "B", "2", headerNames, 3);
                        if (ishdrCreated)
                        {
                            //Write details
                            #region WRITE DETAILS
                            var isDetailsWritten = UtilManager.WriteReportDetailsV1(shtnameNOTMATCHED, reportFileNme, 4, dgSTSMEntries, false);
                            #endregion
                        }
                    }
                    //MATCHED
                    if (msBadgeTotalMatch.BadgeText != "0")
                    {
                        ////Report Header
                        var ishdrCreated = UtilManager.WriteReportHeaderData(shtnameMATCHED, reportFileNme, "KIRK DISTRIBUTORS", "B", "2", headerNames, 3);
                        if (ishdrCreated)
                        {
                            //Write details
                            #region WRITE DETAILS
                            var isDetailsWritten = UtilManager.WriteReportDetailsV1(shtnameMATCHED, reportFileNme, 4, dgSTSMEntries, true);
                            #endregion
                        }
                    }
                    #endregion

                    MetroSetMessageBox.Show(this, "Reports Generated Successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MetroSetMessageBox.Show(this, "Unable to generate report.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MetroSetMessageBox.Show(this, $"{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        public void GenerateNotMatchReportWISYNCO()
        {
            int cnt = 4, lrow = 0, lrow2 = 0;
            decimal totDebitAmt = 0, totCreditAmt = 0, totBalance = 0;
            List<string> headerNames = new List<string> { "TransactionRef", "Date", "Document Type", "Debit Amount", "Credit Amount", "Balance" };
            try
            {
                if (File.Exists(txtStmtFileLoc.Text.Trim()))
                {
                    XLWorkbook workbook = new XLWorkbook(txtStmtFileLoc.Text.Trim());
                    IXLWorksheet ws = workbook.Worksheet("NOTMATCH_WISYNCO");
                    //Clear worksheet
                    ws.Clear();

                    #region WRITE HEADER
                    //Report Header
                    ws.Cell($"B2").Value = "Wisynco Report";
                    ws.Cell($"B2").Style.Fill.SetBackgroundColor(XLColor.Blue);
                    // Set the color for the entire cell
                    ws.Cell($"B2").Style.Font.FontColor = XLColor.White;
                    ws.Cell($"B2").Style.Font.Bold = true;
                    ws.Cell($"B2").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    //Column headings

                    foreach (var txt in headerNames)
                    {
                        int cntHd = headerNames.IndexOf(txt) + 2;
                        ws.Cell(3, cntHd).Value = txt;
                        ws.Cell(3, cntHd).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        ws.Cell(3, cntHd).Style.Fill.SetBackgroundColor(XLColor.Blue);
                        // Set the color for the entire cell
                        ws.Cell(3, cntHd).Style.Font.FontColor = XLColor.White;
                        ws.Cell(3, cntHd).Style.Font.Bold = true;
                    }
                    #endregion
                    foreach (DataGridViewRow r in dgSTSMEntries.Rows)
                    {
                        //Look for unmatched statements
                        if (Convert.ToBoolean(r.Cells["chbIsMatch"].Value) == false)
                        {
                            //Write data to column B
                            ws.Cell($"B{cnt}").Value = r.Cells["ReferenceNo"].Value.ToString();
                            ws.Cell($"B{cnt}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                            //Write data to column C
                            ws.Cell($"C{cnt}").Value = r.Cells["Date"].Value.ToString();
                            //Write data to column D
                            ws.Cell($"D{cnt}").Value = r.Cells["DocType"].Value.ToString();
                            //Write data to column E
                            ws.Cell($"E{cnt}").Value = r.Cells["Debit"].Value.ToString().Replace(",", "");
                            ws.Cell($"E{cnt}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
                            totDebitAmt = totDebitAmt + Convert.ToDecimal(r.Cells["Debit"].Value.ToString().Replace(",", ""));
                            //Write data to column F
                            ws.Cell($"F{cnt}").Value = r.Cells["Credit"].Value.ToString().Replace(",", "");
                            ws.Cell($"F{cnt}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
                            totCreditAmt = totCreditAmt + Convert.ToDecimal(r.Cells["Credit"].Value.ToString().Replace(",", ""));
                            //Write data to column G
                            ws.Cell($"G{cnt}").Value = r.Cells["Balance"].Value.ToString().Replace(",", "");
                            ws.Cell($"G{cnt}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
                            totBalance = totBalance + Convert.ToDecimal(r.Cells["Balance"].Value.ToString().Replace(",", ""));
                            cnt += 1;
                        }
                        //cnt += 1;
                    }
                    //Styling
                    lrow = ws.LastRowUsed().RowNumber();
                    ws.Cell($"E{lrow + 1}").Value = totDebitAmt.ToString();
                    ws.Cell($"E{lrow + 1}").Style.Font.Bold = true;
                    ws.Cell($"E{lrow + 1}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";

                    ws.Cell($"F{lrow + 1}").Value = totCreditAmt.ToString();
                    ws.Cell($"F{lrow + 1}").Style.Font.Bold = true;
                    ws.Cell($"F{lrow + 1}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";

                    ws.Cell($"G{lrow + 1}").Value = totBalance.ToString();
                    ws.Cell($"G{lrow + 1}").Style.Font.Bold = true;
                    ws.Cell($"G{lrow + 1}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";

                    int newRow = lrow + 1;
                    ws.Range($"B{newRow}:G{newRow}").Style.Border.TopBorder = XLBorderStyleValues.Thin;

                    lrow2 = ws.LastRowUsed().RowNumber();

                    IXLRange range = ws.Range(ws.Cell($"B4").Address, ws.Cell($"G{lrow2}").Address);

                    range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                    ws.Columns().AdjustToContents();
                    workbook.SaveAs(txtStmtFileLoc.Text.Trim());
                    //MessageBox.Show("Reports Generated Successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    MetroSetMessageBox.Show(this, "Reports Generated Successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MetroSetMessageBox.Show(this, $"{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        public void LoadVendorDropDownList()
        {
            var jsonFIleFUllPath = AppDomain.CurrentDomain.BaseDirectory + @"\Vendors.json";
            var vendorLst = UtilManager.GetVendors(jsonFIleFUllPath);
            var lstVendors = vendorLst.OrderBy(v => v.Name).ToList();

            //Insert the Default Item to List.
            lstVendors.Insert(0, new Vendor
            {
                Id = "0",
                Name = "Please select vendor"
            });
            //cboVendors.SelectedIndex = 0;

            cboVendors.DataSource = lstVendors;
            cboVendors.DisplayMember = "Name";
            cboVendors.ValueMember = "Id";
        }
        public frmLoadStatement()
        {
            InitializeComponent();
        }

        private void frmLoadStatement_Load(object sender, EventArgs e)
        {
            LoadVendorDropDownList();
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
                cboVendors.Visible = true;
            }
        }

        private void cboVendors_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboVendors.Text.ToLower() != "please select vendor")
                btnLoadEntries.Visible = true;
            else
                btnLoadEntries.Visible = false;
        }

        private void btnLoadEntries_Click(object sender, EventArgs e)
        {
            try
            {
                string vendor = cboVendors.SelectedValue.ToString().Trim();

                gbValResults.Visible = false;
                panelDataSummaryMassy.Visible = false;
                dgSTSMEntries.Visible = false;
                msBadgeTotalMatch.BadgeText = "0";
                msBadgeTotalNotMatch.BadgeText = "0";
                msBadgeTotalRec.BadgeText = "0";
                panelDataSummaryMassy.Visible = false;
                lblTotalAmount.Visible = false;
                lblTotalAmount2.Visible = false;
                lblOriginalAmt.Visible = false;
                lblOriginalAmt.Visible = false;
                lblOutstandingAmt.Visible = false;
                lblOutstandingAmt2.Visible = false;

                (Application.OpenForms["Main"] as Main).slblStatus.Text = "";
                //((Main)this.Owner).slblStatus.Text = "";
                lblProgressPercent.Visible = false;
                progressBar1.Visible = false;
                btnGenerateReports.Visible = false;
                //Load datable
                if (vendor.ToUpper() == "MASSY")
                {
                    var massyEntries = DataAccess.GetSTMTEntriesMassy(vendor.ToUpper(),txtStmtFileLoc.Text.Trim());
                    dgSTSMEntries.DataSource = null;
                    dgSTSMEntries.DataSource = massyEntries;

                    if (massyEntries != null)
                    {
                        var totalAmt = massyEntries.Sum(t => Convert.ToDecimal(t.Amount));


                        lblOutstandingAmt.Text = String.Format("{0:C}", totalAmt);

                        lblOriginalAmt.Visible = false;
                        lblOriginalAmt2.Visible = false;

                        lblOutstandingAmt.Visible = true;
                        lblOutstandingAmt2.Text = "Total Amount:-";
                        lblOutstandingAmt2.Visible = true;

                        btnValidate.Visible = true;
                        gbValResults.Visible = true;
                        panelDataSummaryMassy.Visible = true;
                        dgSTSMEntries.Visible = true;
                        (Application.OpenForms["Main"] as Main).slblStatus.Text = "Template Loaded Successfully!";
                        //var totalOriginalAmt = massyEntries.Sum(t => Convert.ToDecimal(t.OriginalAmount));
                        //var totalOutstandingAmt = massyEntries.Sum(t => Convert.ToDecimal(t.OutstandingAmount));

                        //lblOriginalAmt.Text = String.Format("{0:C}", totalOriginalAmt);
                        //lblOutstandingAmt.Text = String.Format("{0:C}", totalOutstandingAmt);

                        //lblOriginalAmt.Visible = true;
                        //lblOriginalAmt2.Text = "Total Original Amount:-";
                        //lblOriginalAmt2.Visible = true;

                        //lblOutstandingAmt.Visible = true;
                        //lblOutstandingAmt2.Text = "Total Outstanding Amount:-";
                        //lblOutstandingAmt2.Visible = true;

                        //btnValidate.Visible = true;
                        //gbValResults.Visible = true;
                        //panelDataSummaryMassy.Visible = true;
                        //dgSTSMEntries.Visible = true;
                        //(Application.OpenForms["Main"] as Main).slblStatus.Text = "Template Loaded Successfully!";
                    }
                    else
                    {
                        btnValidate.Visible = false;
                        gbValResults.Visible = false;
                        MetroSetMessageBox.Show(this, $"No data found for {cboVendors.Text}!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                }
                else if (vendor.ToUpper() == "WISYNCO")
                {
                    var wisyncoEntries = DataAccess.GetSTMTEntriesWisynco(vendor.ToUpper(), txtStmtFileLoc.Text.Trim());
                    dgSTSMEntries.DataSource = null;
                    dgSTSMEntries.DataSource = wisyncoEntries;

                    if (wisyncoEntries != null)
                    {
                        var totalCreditAmt = wisyncoEntries.Sum(t => Convert.ToDecimal(t.Credit));
                        var totalDebitAmt = wisyncoEntries.Sum(t => Convert.ToDecimal(t.Debit));
                        var totalBalance = wisyncoEntries.Sum(t => Convert.ToDecimal(t.Balance));

                        lblOriginalAmt.Text = String.Format("{0:C}", totalCreditAmt);
                        lblTotalAmount.Text = String.Format("{0:C}", totalDebitAmt);
                        lblOutstandingAmt.Text = String.Format("{0:C}", totalBalance);

                        lblTotalAmount.Visible = true;
                        lblTotalAmount2.Text = "Total Debit:-";
                        lblTotalAmount2.Visible = true;

                        lblOriginalAmt.Visible = true;
                        lblOriginalAmt2.Text = "Total Credit:-";
                        lblOriginalAmt2.Visible = true;

                        lblOutstandingAmt.Visible = true;
                        lblOutstandingAmt2.Text = "Total Balance:-";
                        lblOutstandingAmt2.Visible = true;

                        panelDataSummaryMassy.Visible = true;
                        btnValidate.Visible = true;
                        gbValResults.Visible = true;
                        dgSTSMEntries.Visible = true;
                        (Application.OpenForms["Main"] as Main).slblStatus.Text = "Template Loaded Successfully!";
                    }
                    else
                    {
                        btnValidate.Visible = false;
                        gbValResults.Visible = false;
                        MetroSetMessageBox.Show(this, $"No data found for {cboVendors.Text}!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else if (vendor.ToUpper() == "CARIBPRODUCER")
                {
                    var caribProdEntries = DataAccess.GetSTMTEntriesCaribProducers(vendor.ToUpper(), txtStmtFileLoc.Text.Trim());
                    dgSTSMEntries.DataSource = null;
                    dgSTSMEntries.DataSource = caribProdEntries;

                    if (caribProdEntries != null)
                    {
                        var totalAmt = caribProdEntries.Sum(t => Convert.ToDecimal(t.Amount));
                        var totalBal = caribProdEntries.Sum(t => Convert.ToDecimal(t.Balance.Replace(",", "")));

                        lblOriginalAmt.Text = String.Format("{0:C}", totalAmt);
                        lblOutstandingAmt.Text = String.Format("{0:C}", totalBal);

                        lblOriginalAmt.Visible = true;
                        lblOriginalAmt2.Text = "Total Amount:-";
                        lblOriginalAmt2.Visible = true;

                        lblOutstandingAmt.Visible = true;
                        lblOutstandingAmt2.Text = "Total Balance:-";
                        lblOutstandingAmt2.Visible = true;

                        btnValidate.Visible = true;
                        gbValResults.Visible = true;
                        panelDataSummaryMassy.Visible = true;
                        dgSTSMEntries.Visible = true;
                        (Application.OpenForms["Main"] as Main).slblStatus.Text = "Template Loaded Successfully!";
                    }
                    else
                    {
                        btnValidate.Visible = false;
                        gbValResults.Visible = false;
                        MetroSetMessageBox.Show(this, $"No data found for {cboVendors.Text}!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else if (vendor.ToUpper() == "CONSOLBAKERIES")
                {
                    var consolBakeriesEntries = DataAccess.GetSTMTEntriesConsolBakeries(vendor.ToUpper(), txtStmtFileLoc.Text.Trim());
                    dgSTSMEntries.DataSource = null;
                    dgSTSMEntries.DataSource = consolBakeriesEntries;

                    if (consolBakeriesEntries != null)
                    {
                        var totalAmt = consolBakeriesEntries.Sum(t => Convert.ToDecimal(t.Amount));


                        lblOutstandingAmt.Text = String.Format("{0:C}", totalAmt);

                        lblOriginalAmt.Visible = false;
                        lblOriginalAmt2.Visible = false;

                        lblOutstandingAmt.Visible = true;
                        lblOutstandingAmt2.Text = "Total Amount:-";
                        lblOutstandingAmt2.Visible = true;

                        btnValidate.Visible = true;
                        gbValResults.Visible = true;
                        panelDataSummaryMassy.Visible = true;
                        dgSTSMEntries.Visible = true;
                        (Application.OpenForms["Main"] as Main).slblStatus.Text = "Template Loaded Successfully!";
                    }
                    else
                    {
                        btnValidate.Visible = false;
                        gbValResults.Visible = false;
                        MetroSetMessageBox.Show(this, $"No data found for {cboVendors.Text}!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else if (vendor.ToUpper() == "CONTBAKING")
                {
                    var continentalEntries = DataAccess.GetSTMTEntriesContinentalBaking(vendor.ToUpper(), txtStmtFileLoc.Text.Trim());
                    dgSTSMEntries.DataSource = null;
                    dgSTSMEntries.DataSource = continentalEntries;

                    if (continentalEntries != null)
                    {
                        var totalAmt = continentalEntries.Sum(t => Convert.ToDecimal(t.Amount));


                        lblOutstandingAmt.Text = String.Format("{0:C}", totalAmt);

                        lblOriginalAmt.Visible = false;
                        lblOriginalAmt2.Visible = false;

                        lblOutstandingAmt.Visible = true;
                        lblOutstandingAmt2.Text = "Total Amount:-";
                        lblOutstandingAmt2.Visible = true;

                        btnValidate.Visible = true;
                        gbValResults.Visible = true;
                        panelDataSummaryMassy.Visible = true;
                        dgSTSMEntries.Visible = true;
                        (Application.OpenForms["Main"] as Main).slblStatus.Text = "Template Loaded Successfully!";
                    }
                    else
                    {
                        btnValidate.Visible = false;
                        gbValResults.Visible = false;
                        MetroSetMessageBox.Show(this, $"No data found for {cboVendors.Text}!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else if (vendor.ToUpper() == "RAINFOREST")
                {
                    var rainforestEntries = DataAccess.GetSTMTEntriesRainforest(vendor.ToUpper(), txtStmtFileLoc.Text.Trim());
                    dgSTSMEntries.DataSource = null;
                    dgSTSMEntries.DataSource = rainforestEntries;

                    if (rainforestEntries != null)
                    {
                        var totalInvAmt = rainforestEntries.Sum(t => Convert.ToDecimal(t.InvoiceAmount));
                        var totalPmtAmt = rainforestEntries.Sum(t => Convert.ToDecimal(t.Payments));
                        var totalBalance = rainforestEntries.Sum(t => Convert.ToDecimal(t.Balance));

                        lblOriginalAmt.Text = String.Format("{0:C}", totalPmtAmt);
                        lblTotalAmount.Text = String.Format("{0:C}", totalInvAmt);
                        lblOutstandingAmt.Text = String.Format("{0:C}", totalBalance);

                        lblTotalAmount.Visible = true;
                        lblTotalAmount2.Text = "Total Invoice Amount:-";
                        lblTotalAmount2.Visible = true;

                        lblOriginalAmt.Visible = true;
                        lblOriginalAmt2.Text = "Total Payments:-";
                        lblOriginalAmt2.Visible = true;

                        lblOutstandingAmt.Visible = true;
                        lblOutstandingAmt2.Text = "Total Balance:-";
                        lblOutstandingAmt2.Visible = true;

                        panelDataSummaryMassy.Visible = true;
                        btnValidate.Visible = true;
                        gbValResults.Visible = true;
                        dgSTSMEntries.Visible = true;
                        (Application.OpenForms["Main"] as Main).slblStatus.Text = "Template Loaded Successfully!";
                    }
                    else
                    {
                        dgSTSMEntries.DataSource = null;
                        btnValidate.Visible = false;
                        gbValResults.Visible = false;
                        MetroSetMessageBox.Show(this, $"No data found for {cboVendors.Text}!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else if (vendor.ToUpper() == "WORLDBRANDS")
                {
                    var wbEntries = DataAccess.GetSTMTEntriesWorldBrands(vendor.ToUpper(), txtStmtFileLoc.Text.Trim());
                    dgSTSMEntries.DataSource = null;
                    dgSTSMEntries.DataSource = wbEntries;

                    if (wbEntries != null)
                    {
                        var totalAmt = wbEntries.Sum(t => Convert.ToDecimal(t.Amount));


                        lblOutstandingAmt.Text = String.Format("{0:C}", totalAmt);

                        lblOriginalAmt.Visible = false;
                        lblOriginalAmt2.Visible = false;

                        lblOutstandingAmt.Visible = true;
                        lblOutstandingAmt2.Text = "Total Amount:-";
                        lblOutstandingAmt2.Visible = true;

                        btnValidate.Visible = true;
                        gbValResults.Visible = true;
                        panelDataSummaryMassy.Visible = true;
                        dgSTSMEntries.Visible = true;
                        (Application.OpenForms["Main"] as Main).slblStatus.Text = "Template Loaded Successfully!";
                    }
                    else
                    {
                        btnValidate.Visible = false;
                        gbValResults.Visible = false;
                        MetroSetMessageBox.Show(this, $"No data found for {cboVendors.Text}!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else if (vendor.ToUpper() == "CONSBRANDS")
                {
                    var cbEntries = DataAccess.GetSTMTEntriesConsumerBrands(vendor.ToUpper(), txtStmtFileLoc.Text.Trim());
                    dgSTSMEntries.DataSource = null;
                    dgSTSMEntries.DataSource = cbEntries;

                    if (cbEntries != null)
                    {
                        var totalAmt = cbEntries.Sum(t => Convert.ToDecimal(t.Amount));


                        lblOutstandingAmt.Text = String.Format("{0:C}", totalAmt);

                        lblOriginalAmt.Visible = false;
                        lblOriginalAmt2.Visible = false;

                        lblOutstandingAmt.Visible = true;
                        lblOutstandingAmt2.Text = "Total Amount:-";
                        lblOutstandingAmt2.Visible = true;

                        btnValidate.Visible = true;
                        gbValResults.Visible = true;
                        panelDataSummaryMassy.Visible = true;
                        dgSTSMEntries.Visible = true;
                        (Application.OpenForms["Main"] as Main).slblStatus.Text = "Template Loaded Successfully!";
                    }
                    else
                    {
                        btnValidate.Visible = false;
                        gbValResults.Visible = false;
                        MetroSetMessageBox.Show(this, $"No data found for {cboVendors.Text}!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else if (vendor.ToUpper() == "FACEYPHARMACY")
                {
                    var fpEntries = DataAccess.GetSTMTEntriesFaceyPharmacy(vendor.ToUpper(), txtStmtFileLoc.Text.Trim());
                    dgSTSMEntries.DataSource = null;
                    dgSTSMEntries.DataSource = fpEntries;

                    if (fpEntries != null)
                    {
                        var totalAmt = fpEntries.Sum(t => Convert.ToDecimal(t.Amount));
                        var totalBal = fpEntries.Sum(t => Convert.ToDecimal(t.Balance));

                        lblOutstandingAmt.Text = String.Format("{0:C}", totalBal);
                        lblOriginalAmt.Text = String.Format("{0:C}", totalAmt);

                        lblOriginalAmt.Visible = false;
                        lblOriginalAmt2.Visible = false;

                        lblOriginalAmt.Visible = true;
                        lblOriginalAmt2.Text = "Total Amount:-";
                        lblOriginalAmt2.Visible = true;

                        lblOutstandingAmt.Visible = false;
                        lblOutstandingAmt2.Visible = false;

                        lblOutstandingAmt.Visible = true;
                        lblOutstandingAmt2.Text = "Total Balance:-";
                        lblOutstandingAmt2.Visible = true;

                        btnValidate.Visible = true;
                        gbValResults.Visible = true;
                        panelDataSummaryMassy.Visible = true;
                        dgSTSMEntries.Visible = true;
                        (Application.OpenForms["Main"] as Main).slblStatus.Text = "Template Loaded Successfully!";
                    }
                    else
                    {
                        btnValidate.Visible = false;
                        gbValResults.Visible = false;
                        MetroSetMessageBox.Show(this, $"No data found for {cboVendors.Text}!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else if (vendor.ToUpper() == "CARIMED")
                {
                    var carimedEntries = DataAccess.GetSTMTEntriesCarimed(vendor.ToUpper(), txtStmtFileLoc.Text.Trim());
                    dgSTSMEntries.DataSource = null;
                    dgSTSMEntries.DataSource = carimedEntries;
                    if (carimedEntries != null)
                    {
                        var totalAmt = carimedEntries.Sum(t => Convert.ToDecimal(t.Amount));
                        var totalBal = carimedEntries.Sum(t => Convert.ToDecimal(t.Balance));

                        lblOutstandingAmt.Text = String.Format("{0:C}", totalBal);
                        lblOriginalAmt.Text = String.Format("{0:C}", totalAmt);

                        lblOriginalAmt.Visible = false;
                        lblOriginalAmt2.Visible = false;

                        lblOriginalAmt.Visible = true;
                        lblOriginalAmt2.Text = "Total Amount:-";
                        lblOriginalAmt2.Visible = true;

                        lblOutstandingAmt.Visible = false;
                        lblOutstandingAmt2.Visible = false;

                        lblOutstandingAmt.Visible = true;
                        lblOutstandingAmt2.Text = "Total Balance:-";
                        lblOutstandingAmt2.Visible = true;

                        btnValidate.Visible = true;
                        gbValResults.Visible = true;
                        panelDataSummaryMassy.Visible = true;
                        dgSTSMEntries.Visible = true;
                        (Application.OpenForms["Main"] as Main).slblStatus.Text = "Template Loaded Successfully!";
                    }
                    else
                    {
                        btnValidate.Visible = false;
                        gbValResults.Visible = false;
                        MetroSetMessageBox.Show(this, $"No data found for {cboVendors.Text}!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else if (vendor.ToUpper() == "FACEY")
                {
                    var faceyEntries = DataAccess.GetSTMTEntriesFacey(vendor.ToUpper(), txtStmtFileLoc.Text.Trim());
                    dgSTSMEntries.DataSource = null;
                    dgSTSMEntries.DataSource = faceyEntries;
                    if (faceyEntries != null)
                    {
                        var totalAmt = faceyEntries.Sum(t => Convert.ToDecimal(t.Amount));
                        var totalBal = faceyEntries.Sum(t => Convert.ToDecimal(t.Balance));

                        lblOutstandingAmt.Text = String.Format("{0:C}", totalBal);
                        lblOriginalAmt.Text = String.Format("{0:C}", totalAmt);

                        lblOriginalAmt.Visible = false;
                        lblOriginalAmt2.Visible = false;

                        lblOriginalAmt.Visible = true;
                        lblOriginalAmt2.Text = "Total Amount:-";
                        lblOriginalAmt2.Visible = true;

                        lblOutstandingAmt.Visible = false;
                        lblOutstandingAmt2.Visible = false;

                        lblOutstandingAmt.Visible = true;
                        lblOutstandingAmt2.Text = "Total Balance:-";
                        lblOutstandingAmt2.Visible = true;

                        btnValidate.Visible = true;
                        gbValResults.Visible = true;
                        panelDataSummaryMassy.Visible = true;
                        dgSTSMEntries.Visible = true;
                        (Application.OpenForms["Main"] as Main).slblStatus.Text = "Template Loaded Successfully!";
                    }
                    else
                    {
                        btnValidate.Visible = false;
                        gbValResults.Visible = false;
                        MetroSetMessageBox.Show(this, $"No data found for {cboVendors.Text}!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else if (vendor.ToUpper() == "DERRIMON")
                {
                    var derrimonEntries = DataAccess.GetSTMTEntriesDerrimon(vendor.ToUpper(), txtStmtFileLoc.Text.Trim());
                    dgSTSMEntries.DataSource = null;
                    dgSTSMEntries.DataSource = derrimonEntries;

                    if (derrimonEntries != null)
                    {
                        var totalAmt = derrimonEntries.Sum(t => Convert.ToDecimal(t.Amount));


                        lblOutstandingAmt.Text = String.Format("{0:C}", totalAmt);

                        lblOriginalAmt.Visible = false;
                        lblOriginalAmt2.Visible = false;

                        lblOutstandingAmt.Visible = true;
                        lblOutstandingAmt2.Text = "Total Amount:-";
                        lblOutstandingAmt2.Visible = true;

                        btnValidate.Visible = true;
                        gbValResults.Visible = true;
                        panelDataSummaryMassy.Visible = true;
                        dgSTSMEntries.Visible = true;
                        (Application.OpenForms["Main"] as Main).slblStatus.Text = "Template Loaded Successfully!";
                    }
                    else
                    {
                        btnValidate.Visible = false;
                        gbValResults.Visible = false;
                        MetroSetMessageBox.Show(this, $"No data found for {cboVendors.Text}!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else if (vendor.ToUpper() == "BESTDRESSED")
                {
                    var bdEntries = DataAccess.GetSTMTEntriesBestDressed(vendor.ToUpper(), txtStmtFileLoc.Text.Trim());
                    dgSTSMEntries.DataSource = null;
                    dgSTSMEntries.DataSource = bdEntries;

                    if (bdEntries != null)
                    {
                        var totalCreditAmt = bdEntries.Sum(t => Convert.ToDecimal(t.Credit));
                        var totalDebitAmt = bdEntries.Sum(t => Convert.ToDecimal(t.Debit));
                        var totalBalance = bdEntries.Sum(t => Convert.ToDecimal(t.Balance));

                        lblOriginalAmt.Text = String.Format("{0:C}", totalCreditAmt);
                        lblTotalAmount.Text = String.Format("{0:C}", totalDebitAmt);
                        lblOutstandingAmt.Text = String.Format("{0:C}", totalBalance);

                        lblTotalAmount.Visible = true;
                        lblTotalAmount2.Text = "Total Debit:-";
                        lblTotalAmount2.Visible = true;

                        lblOriginalAmt.Visible = true;
                        lblOriginalAmt2.Text = "Total Credit:-";
                        lblOriginalAmt2.Visible = true;

                        lblOutstandingAmt.Visible = true;
                        lblOutstandingAmt2.Text = "Total Balance:-";
                        lblOutstandingAmt2.Visible = true;

                        panelDataSummaryMassy.Visible = true;
                        btnValidate.Visible = true;
                        gbValResults.Visible = true;
                        dgSTSMEntries.Visible = true;
                        (Application.OpenForms["Main"] as Main).slblStatus.Text = "Template Loaded Successfully!";
                    }
                    else
                    {
                        btnValidate.Visible = false;
                        gbValResults.Visible = false;
                        MetroSetMessageBox.Show(this, $"No data found for {cboVendors.Text}!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else if (vendor.ToUpper() == "COPPERWOOD")
                {
                    var cwEntries = DataAccess.GetSTMTEntriesCopperwood(vendor.ToUpper(), txtStmtFileLoc.Text.Trim());
                    dgSTSMEntries.DataSource = null;
                    dgSTSMEntries.DataSource = cwEntries;

                    if (cwEntries != null)
                    {
                        var totalAmt = cwEntries.Sum(t => Convert.ToDecimal(t.Amount));
                        var totalBal = cwEntries.Sum(t => Convert.ToDecimal(t.Balance));

                        lblOutstandingAmt.Text = String.Format("{0:C}", totalBal);
                        lblOriginalAmt.Text = String.Format("{0:C}", totalAmt);

                        lblOriginalAmt.Visible = false;
                        lblOriginalAmt2.Visible = false;

                        lblOriginalAmt.Visible = true;
                        lblOriginalAmt2.Text = "Total Amount:-";
                        lblOriginalAmt2.Visible = true;

                        lblOutstandingAmt.Visible = false;
                        lblOutstandingAmt2.Visible = false;

                        lblOutstandingAmt.Visible = true;
                        lblOutstandingAmt2.Text = "Total Balance:-";
                        lblOutstandingAmt2.Visible = true;

                        btnValidate.Visible = true;
                        gbValResults.Visible = true;
                        panelDataSummaryMassy.Visible = true;
                        dgSTSMEntries.Visible = true;
                        (Application.OpenForms["Main"] as Main).slblStatus.Text = "Template Loaded Successfully!";
                    }
                    else
                    {
                        btnValidate.Visible = false;
                        gbValResults.Visible = false;
                        MetroSetMessageBox.Show(this, $"No data found for {cboVendors.Text}!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else if (vendor.ToUpper() == "TGEDDES")
                {
                    var tgeddesEntries = DataAccess.GetSTMTEntriesTGeddes(vendor.ToUpper(), txtStmtFileLoc.Text.Trim());
                    dgSTSMEntries.DataSource = null;
                    dgSTSMEntries.DataSource = tgeddesEntries;

                    if (tgeddesEntries != null)
                    {
                        var totalAmt = tgeddesEntries.Sum(t => Convert.ToDecimal(t.Amount));
                        var totalBal = tgeddesEntries.Sum(t => Convert.ToDecimal(t.Balance));

                        lblOutstandingAmt.Text = String.Format("{0:C}", totalBal);
                        lblOriginalAmt.Text = String.Format("{0:C}", totalAmt);

                        lblOriginalAmt.Visible = false;
                        lblOriginalAmt2.Visible = false;

                        lblOriginalAmt.Visible = true;
                        lblOriginalAmt2.Text = "Total Amount:-";
                        lblOriginalAmt2.Visible = true;

                        lblOutstandingAmt.Visible = false;
                        lblOutstandingAmt2.Visible = false;

                        lblOutstandingAmt.Visible = true;
                        lblOutstandingAmt2.Text = "Total Balance:-";
                        lblOutstandingAmt2.Visible = true;

                        btnValidate.Visible = true;
                        gbValResults.Visible = true;
                        panelDataSummaryMassy.Visible = true;
                        dgSTSMEntries.Visible = true;
                        (Application.OpenForms["Main"] as Main).slblStatus.Text = "Template Loaded Successfully!";
                    }
                    else
                    {
                        btnValidate.Visible = false;
                        gbValResults.Visible = false;
                        MetroSetMessageBox.Show(this, $"No data found for {cboVendors.Text}!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else if (vendor.ToUpper() == "CHASERAMSON")
                {
                    var chasEntries = DataAccess.GetSTMTEntriesChasERamson(vendor.ToUpper(), txtStmtFileLoc.Text.Trim());
                    dgSTSMEntries.DataSource = null;
                    dgSTSMEntries.DataSource = chasEntries;

                    if (chasEntries != null)
                    {
                        var totalCreditAmt = chasEntries.Sum(t => Convert.ToDecimal(t.Credit));
                        var totalDebitAmt = chasEntries.Sum(t => Convert.ToDecimal(t.Debit));
                        var totalBalance = chasEntries.Sum(t => Convert.ToDecimal(t.Balance));

                        lblOriginalAmt.Text = String.Format("{0:C}", totalCreditAmt);
                        lblTotalAmount.Text = String.Format("{0:C}", totalDebitAmt);
                        lblOutstandingAmt.Text = String.Format("{0:C}", totalBalance);

                        lblTotalAmount.Visible = true;
                        lblTotalAmount2.Text = "Total Debit:-";
                        lblTotalAmount2.Visible = true;

                        lblOriginalAmt.Visible = true;
                        lblOriginalAmt2.Text = "Total Credit:-";
                        lblOriginalAmt2.Visible = true;

                        lblOutstandingAmt.Visible = true;
                        lblOutstandingAmt2.Text = "Total Balance:-";
                        lblOutstandingAmt2.Visible = true;

                        panelDataSummaryMassy.Visible = true;
                        btnValidate.Visible = true;
                        gbValResults.Visible = true;
                        dgSTSMEntries.Visible = true;
                        (Application.OpenForms["Main"] as Main).slblStatus.Text = "Template Loaded Successfully!";
                    }
                    else
                    {
                        btnValidate.Visible = false;
                        gbValResults.Visible = false;
                        MetroSetMessageBox.Show(this, $"No data found for {cboVendors.Text}!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else if (vendor.ToUpper() == "CELBRANDS")
                {
                    //var celebrandsEntries = GetSTMTEntriesCeleBrands(vendor.ToUpper());
                    //dgSTSMEntries.DataSource = null;
                    //dgSTSMEntries.DataSource = celebrandsEntries;

                    //if (celebrandsEntries != null)
                    //{
                    //    var totalCreditAmt = celebrandsEntries.Sum(t => Convert.ToDecimal(t.Credit));
                    //    var totalDebitAmt = celebrandsEntries.Sum(t => Convert.ToDecimal(t.Debit));
                    //    var totalBalance = celebrandsEntries.Sum(t => Convert.ToDecimal(t.Balance));

                    //    lblOriginalAmt.Text = String.Format("{0:C}", totalCreditAmt);
                    //    lblTotalAmount.Text = String.Format("{0:C}", totalDebitAmt);
                    //    lblOutstandingAmt.Text = String.Format("{0:C}", totalBalance);

                    //    lblTotalAmount.Visible = true;
                    //    lblTotalAmount2.Text = "Total Debit:-";
                    //    lblTotalAmount2.Visible = true;

                    //    lblOriginalAmt.Visible = true;
                    //    lblOriginalAmt2.Text = "Total Credit:-";
                    //    lblOriginalAmt2.Visible = true;

                    //    lblOutstandingAmt.Visible = true;
                    //    lblOutstandingAmt2.Text = "Total Balance:-";
                    //    lblOutstandingAmt2.Visible = true;

                    //    panelDataSummaryMassy.Visible = true;
                    //    btnValidate.Visible = true;
                    //    gbValResults.Visible = true;
                    //    dgSTSMEntries.Visible = true;
                    //    (Application.OpenForms["Main"] as Main).slblStatus.Text = "Template Loaded Successfully!";
                    //}
                    //else
                    //{
                    //    btnValidate.Visible = false;
                    //    gbValResults.Visible = false;
                    //    MetroSetMessageBox.Show(this, $"No data found for {cboVendors.Text}!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    //}
                }
                else if (vendor.ToUpper() == "KIRK")
                {
                    var kirkEntries = DataAccess.GetSTMTEntriesKIRKDISTRIBUTORS(vendor.ToUpper(), txtStmtFileLoc.Text.Trim());
                    dgSTSMEntries.DataSource = null;
                    dgSTSMEntries.DataSource = kirkEntries;

                    if (kirkEntries != null)
                    {
                        var totalAmt = kirkEntries.Sum(t => Convert.ToDecimal(t.Amount));


                        lblOutstandingAmt.Text = String.Format("{0:C}", totalAmt);

                        lblOriginalAmt.Visible = false;
                        lblOriginalAmt2.Visible = false;

                        lblOutstandingAmt.Visible = true;
                        lblOutstandingAmt2.Text = "Total Amount:-";
                        lblOutstandingAmt2.Visible = true;

                        btnValidate.Visible = true;
                        gbValResults.Visible = true;
                        panelDataSummaryMassy.Visible = true;
                        dgSTSMEntries.Visible = true;
                        (Application.OpenForms["Main"] as Main).slblStatus.Text = "Template Loaded Successfully!";
                    }
                    else
                    {
                        btnValidate.Visible = false;
                        gbValResults.Visible = false;
                        MetroSetMessageBox.Show(this, $"No data found for {cboVendors.Text}!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }

                else
                {
                    dgSTSMEntries.DataSource = null;
                    btnValidate.Visible = false;
                    gbValResults.Visible = false;
                    MetroSetMessageBox.Show(this, $"No data found for {cboVendors.Text}!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                dgSTSMEntries.AutoSizeColumnsMode =
                DataGridViewAutoSizeColumnsMode.Fill;

                dgSTSMEntries.AutoResizeColumns();

                if(dgSTSMEntries.DataSource != null)
                    msBadgeTotalRec.BadgeText = dgSTSMEntries.Rows.Count.ToString();
                else
                    msBadgeTotalRec.BadgeText = "0";
            }
            catch (Exception ex)
            {
                MetroSetMessageBox.Show(this, $"{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnValidate_Click(object sender, EventArgs e)
        {
            vendor = "";
            vendor = cboVendors.SelectedValue.ToString().Trim().ToUpper();
            try
            {             
                if (vendor == "MASSY")
                {
                    var sapmaster = DataAccess.GetSAPMasterData($"{vendor}_SAPMASTER", txtStmtFileLoc.Text.Trim());

                    if (sapmaster.Count() != 0)
                    {
                        ValidateStatement(sapmaster,vendor);
                    }
                    else
                    {
                        MetroSetMessageBox.Show(this, $"No SAP data found in sheet name '{vendor}_SAPMASTER'!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                   
                }
                else if (vendor == "DERRIMON")
                {
                    var sapmaster = DataAccess.GetSAPMasterData($"{vendor}_SAPMASTER", txtStmtFileLoc.Text.Trim());

                    if (sapmaster.Count() != 0)
                    {
                        ValidateStatement(sapmaster, vendor);
                    }
                    else
                    {
                        MetroSetMessageBox.Show(this, $"No SAP data found in sheet name '{vendor}_SAPMASTER'!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }
                else if (vendor == "KIRK")
                {
                    var sapmaster = DataAccess.GetSAPMasterData($"{vendor}_SAPMASTER", txtStmtFileLoc.Text.Trim());

                    if (sapmaster.Count() != 0)
                    {
                        ValidateStatement(sapmaster, vendor);
                    }
                    else
                    {
                        MetroSetMessageBox.Show(this, $"No SAP data found in sheet name '{vendor}_SAPMASTER'!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }
                else if (vendor == "CONSOLBAKERIES")
                {
                    var sapmaster = DataAccess.GetSAPMasterData($"{vendor}_SAPMASTER", txtStmtFileLoc.Text.Trim());

                    if (sapmaster.Count() != 0)
                    {
                        ValidateStatement(sapmaster, vendor);
                    }
                    else
                    {
                        MetroSetMessageBox.Show(this, $"No SAP data found in sheet name '{vendor}_MASTER'!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else if (vendor == "COPPERWOOD")
                {
                    var sapmaster = DataAccess.GetSAPMasterData($"{vendor}_SAPMASTER", txtStmtFileLoc.Text.Trim());

                    if (sapmaster.Count() != 0)
                    {
                        ValidateStatement(sapmaster, vendor);
                    }
                    else
                    {
                        MetroSetMessageBox.Show(this, $"No SAP data found in sheet name '{vendor}_MASTER'!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else if (vendor == "TGEDDES")
                {
                    var sapmaster = DataAccess.GetSAPMasterData($"{vendor}_SAPMASTER", txtStmtFileLoc.Text.Trim());

                    if (sapmaster.Count() != 0)
                    {
                        ValidateStatement(sapmaster, vendor);
                    }
                    else
                    {
                        MetroSetMessageBox.Show(this, $"No SAP data found in sheet name '{vendor}_MASTER'!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else if (vendor == "FACEY")
                {
                    var sapmaster = DataAccess.GetSAPMasterData($"{vendor}_SAPMASTER", txtStmtFileLoc.Text.Trim());

                    if (sapmaster.Count() != 0)
                    {
                        ValidateStatement(sapmaster, vendor);
                    }
                    else
                    {
                        MetroSetMessageBox.Show(this, $"No SAP data found in sheet name '{vendor}_MASTER'!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else if (vendor == "WISYNCO")
                {
                    //var sapmaster = DataAccess.GetSAPMasterData($"{vendor.ToUpper()}_SAPMASTER",txtStmtFileLoc.Text.Trim());
                    //if(sapmaster != null)
                    //{
                    //    ValidateWISYNCO(sapmaster);
                    //}
                    //else
                    //{
                    //    MetroSetMessageBox.Show(this, $"No SAP data found in sheet name '{vendor.ToUpper()}_MASTER'!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //}
                }
                else { }
            }
            catch (Exception ex)
            {
                MetroSetMessageBox.Show(this, $"{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnGenerateReports_Click(object sender, EventArgs e)
        {

            (Application.OpenForms["Main"] as Main).slblStatus.Text = "Generating Reports...!";

            if (vendor.ToUpper() == "MASSY")
                GenerateReportMASSY(vendor);
            else if (vendor.ToUpper() == "CONSOLBAKERIES")
                GenerateReportCONSOLBAKERIES(vendor);
            else if (vendor.ToUpper() == "DERRIMON")
                GenerateReportDERRIMON(vendor);
            else if (vendor.ToUpper() == "COPPERWOOD")
                GenerateReportCOPPERWOOD(vendor);
            else if (vendor.ToUpper() == "TGEDDES")
                GenerateReportTGEDDES(vendor);
            else if (vendor.ToUpper() == "FACEY")
                GenerateReportFACEY(vendor);
            else if (vendor.ToUpper() == "KIRK")
                GenerateReportKIRK(vendor);
            //else if (vendor.ToUpper() == "WISYNCO")
            //    GenerateNotMatchReportWISYNCO();

            (Application.OpenForms["Main"] as Main).slblStatus.Text = "Generate Reports Complete!";
        }


    }
}

//public void GenerateNotMatchReportMASSY(string vendor)
//{
//    int cnt = 4, lrow = 0, lrow2 = 0;
//    decimal totOrigAmt = 0, totOutstandingAmt = 0;
//    List<string> headerNames = new List<string> { "Transaction Reference", "Due Date", "Document Type", "Original Amount", "Outstanding Amount" };
//    try
//    {
//        string reportsDIR = Path.GetDirectoryName(txtStmtFileLoc.Text.Trim()) + $@"\REPORTS\{vendor}";
//        //Create directoy if doesn't exist
//        if (!Directory.Exists(reportsDIR))
//            Directory.CreateDirectory(reportsDIR);

//        string reportFileNme = reportsDIR + $@"\NOTMATCHED_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";

//        string shtnameNOTMATCHED = "NOTMATCHED";
//        string shtnameMATCHED = "MATCHED";
//        //create workbook
//        var isCreated = UtilManager.CreateReportsWorkbook(reportFileNme, shtnameNOTMATCHED, shtnameMATCHED);
//        if (isCreated)
//        {
//            XLWorkbook workbook = new XLWorkbook(reportFileNme);
//            IXLWorksheet ws = workbook.Worksheet(vendor);
//            //Clear worksheet
//            ws.Clear();

//            #region WRITE HEADER
//            //Report Header
//            ws.Cell($"B2").Value = "MASSY DISTRIBUTION";
//            ws.Cell($"B2").Style.Fill.SetBackgroundColor(XLColor.Blue);
//            // Set the color for the entire cell
//            ws.Cell($"B2").Style.Font.FontColor = XLColor.White;
//            ws.Cell($"B2").Style.Font.Bold = true;
//            ws.Cell($"B2").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
//            //Column headings

//            foreach (var txt in headerNames)
//            {
//                int cntHd = headerNames.IndexOf(txt) + 2;
//                ws.Cell(3, cntHd).Value = txt;
//                ws.Cell(3, cntHd).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
//                ws.Cell(3, cntHd).Style.Fill.SetBackgroundColor(XLColor.Blue);
//                // Set the color for the entire cell
//                ws.Cell(3, cntHd).Style.Font.FontColor = XLColor.White;
//                ws.Cell(3, cntHd).Style.Font.Bold = true;
//            }
//            #endregion

//            #region WRITE DETAILS
//            foreach (DataGridViewRow r in dgSTSMEntries.Rows)
//            {
//                //Look for unmatched statements
//                if (Convert.ToBoolean(r.Cells["chbIsMatch"].Value) == false)
//                {
//                    //Write data to column B
//                    ws.Cell($"B{cnt}").Value = r.Cells["TransNo"].Value.ToString();
//                    //Write data to column C
//                    ws.Cell($"C{cnt}").Value = r.Cells["DueDate"].Value.ToString();
//                    //Write data to column D
//                    ws.Cell($"D{cnt}").Value = r.Cells["DocType"].Value.ToString();
//                    //Write data to column E
//                    ws.Cell($"E{cnt}").Value = r.Cells["OriginalAmount"].Value.ToString().Replace(",", "");
//                    ws.Cell($"E{cnt}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
//                    totOrigAmt = totOrigAmt + Convert.ToDecimal(r.Cells["OriginalAmount"].Value.ToString().Replace(",", ""));
//                    //Write data to column F
//                    ws.Cell($"F{cnt}").Value = r.Cells["OutstandingAmount"].Value.ToString().Replace(",", "");
//                    ws.Cell($"F{cnt}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
//                    totOutstandingAmt = totOutstandingAmt + Convert.ToDecimal(r.Cells["OutstandingAmount"].Value.ToString().Replace(",", ""));

//                    cnt += 1;
//                }
//            }
//            #endregion

//            #region STYLE REPORT
//            //Styling
//            lrow = ws.LastRowUsed().RowNumber();
//            ws.Cell($"E{lrow + 1}").Value = totOrigAmt.ToString();
//            ws.Cell($"E{lrow + 1}").Style.Font.Bold = true;
//            ws.Cell($"E{lrow + 1}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
//            ws.Cell($"F{lrow + 1}").Value = totOutstandingAmt.ToString();
//            ws.Cell($"F{lrow + 1}").Style.Font.Bold = true;
//            ws.Cell($"F{lrow + 1}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
//            int newRow = lrow + 1;
//            ws.Range($"B{newRow}:F{newRow}").Style.Border.TopBorder = XLBorderStyleValues.Thin;

//            lrow2 = ws.LastRowUsed().RowNumber();

//            IXLRange range = ws.Range(ws.Cell($"B4").Address, ws.Cell($"F{lrow2}").Address);

//            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

//            ws.Columns().AdjustToContents();
//            #endregion

//            workbook.SaveAs(reportFileNme);
//        }


//        MetroSetMessageBox.Show(this, "Reports Generated Successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
//    }
//    catch (Exception ex)
//    {
//        MetroSetMessageBox.Show(this, $"{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
//        //MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
//    }
//}
//private Task ValidateSTMTEntriesMASSY(IProgress<ProgressReport> progress, List<SAPMaster> SAPMasterData)
//{
//    int cntMatch = 0, cntNotMatch = 0, cnt = 1, totalProcess = dgSTSMEntries.Rows.Count;
//    var progressReport = new ProgressReport();

//    dgSTSMEntries.ClearSelection();

//    try
//    {
//        return Task.Run(() =>
//        {
//            foreach (DataGridViewRow row in dgSTSMEntries.Rows)
//            {
//                progressReport.PercentageComplete = cnt++ * 100 / totalProcess;

//                string docNo = row.Cells["TransNo"].Value.ToString();
//                decimal amount = Convert.ToDecimal(row.Cells["OutstandingAmount"].Value.ToString());
//                DateTime docDate = Convert.ToDateTime(row.Cells["TransDate"].Value.ToString());

//                if (SAPMasterData.Any(s => s.Reference == docNo && s.Amount == amount.ToString() && s.DocumentDate.Value.Date == docDate.Date))
//                {
//                    row.Cells["chbIsMatch"].Value = CheckState.Checked;
//                    row.Cells["chbIsMatch"].Style.BackColor = System.Drawing.Color.LightGreen;
//                    cntMatch += 1;
//                }
//                else
//                {
//                    row.Cells["chbIsMatch"].Value = CheckState.Unchecked;
//                    row.Cells["chbIsMatch"].Style.BackColor = Color.MistyRose;
//                    cntNotMatch += 1;
//                }

//                progressReport.TotalRecords = dgSTSMEntries.Rows.Count;
//                progressReport.TotalMatched = cntMatch;
//                progressReport.TotalNotMatched = cntNotMatch;

//                progress.Report(progressReport);

//            }



//        });
//    }
//    catch (ThreadInterruptedException e)
//    {
//        MetroSetMessageBox.Show(this, $"{e.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
//        //MessageBox.Show(e.Message);
//        return null;
//    }
//}

//private Task ValidateSTMTEntriesConsolBakeries(IProgress<ProgressReport> progress, List<SAPMaster> SAPMasterData)
//{
//    int cntMatch = 0, cntNotMatch = 0, cnt = 1, totalProcess = dgSTSMEntries.Rows.Count;
//    var progressReport = new ProgressReport();

//    dgSTSMEntries.ClearSelection();

//    try
//    {
//        return Task.Run(() =>
//        {
//            foreach (DataGridViewRow row in dgSTSMEntries.Rows)
//            {
//                progressReport.PercentageComplete = cnt++ * 100 / totalProcess;

//                string docNo = row.Cells["ReferenceNo"].Value.ToString();
//                decimal amount = Convert.ToDecimal(row.Cells["Amount"].Value.ToString());
//                DateTime docDate = Convert.ToDateTime(row.Cells["Date"].Value.ToString());

//                if (SAPMasterData.Any(s => s.Reference == docNo && s.Amount == amount.ToString() && s.DocumentDate.Value.Date == docDate.Date))
//                {
//                    row.Cells["chbIsMatch"].Value = CheckState.Checked;
//                    row.Cells["chbIsMatch"].Style.BackColor = System.Drawing.Color.LightGreen;
//                    cntMatch += 1;
//                }
//                else
//                {
//                    row.Cells["chbIsMatch"].Value = CheckState.Unchecked;
//                    row.Cells["chbIsMatch"].Style.BackColor = Color.MistyRose;
//                    cntNotMatch += 1;
//                }

//                progressReport.TotalRecords = dgSTSMEntries.Rows.Count;
//                progressReport.TotalMatched = cntMatch;
//                progressReport.TotalNotMatched = cntNotMatch;

//                progress.Report(progressReport);
//                //Thread.Sleep(1000);
//            }



//        });
//    }
//    catch (ThreadInterruptedException e)
//    {
//        MetroSetMessageBox.Show(this, $"{e.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
//        //MessageBox.Show(e.Message);
//        return null;
//    }
//}

//public void GenerateNotMatchReportCONSOLBAKERIES(string vendor)
//{
//    int cnt = 4, lrow = 0, lrow2 = 0;
//    decimal totAmt = 0;
//    List<string> headerNames = new List<string> { "Transaction Reference", "Document Date", "Document Type", "Amount" };
//    try
//    {

//        string reportsDIR = Path.GetDirectoryName(txtStmtFileLoc.Text.Trim()) + $@"\REPORTS\{vendor}";
//        //Create directoy if doesn't exist
//        if (!Directory.Exists(reportsDIR))
//            Directory.CreateDirectory(reportsDIR);

//        string reportFileNme = reportsDIR + $@"\NOTMATCHED_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";

//        string shtnameNOTMATCHED = "NOTMATCHED";
//        string shtnameMATCHED = "MATCHED";
//        //create workbook
//        var isCreated = UtilManager.CreateReportsWorkbook(reportFileNme, shtnameNOTMATCHED, shtnameMATCHED);
//        if (isCreated)
//        {
//            XLWorkbook workbook = new XLWorkbook(reportFileNme);
//            IXLWorksheet ws = workbook.Worksheet(vendor);
//            //Clear worksheet
//            ws.Clear();

//            #region WRITE HEADER
//            //Report Header
//            ws.Cell($"B2").Value = "CONSOLIDATED BAKERIES (JA) LTD";
//            ws.Cell($"B2").Style.Fill.SetBackgroundColor(XLColor.Blue);
//            // Set the color for the entire cell
//            ws.Cell($"B2").Style.Font.FontColor = XLColor.White;
//            ws.Cell($"B2").Style.Font.Bold = true;
//            ws.Cell($"B2").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
//            //Column headings

//            foreach (var txt in headerNames)
//            {
//                int cntHd = headerNames.IndexOf(txt) + 2;
//                ws.Cell(3, cntHd).Value = txt;
//                ws.Cell(3, cntHd).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
//                ws.Cell(3, cntHd).Style.Fill.SetBackgroundColor(XLColor.Blue);
//                // Set the color for the entire cell
//                ws.Cell(3, cntHd).Style.Font.FontColor = XLColor.White;
//                ws.Cell(3, cntHd).Style.Font.Bold = true;
//            }
//            #endregion

//            #region WRITE DETAILS
//            foreach (DataGridViewRow r in dgSTSMEntries.Rows)
//            {
//                //Look for unmatched statements
//                if (Convert.ToBoolean(r.Cells["chbIsMatch"].Value) == false)
//                {
//                    //Write data to column B
//                    ws.Cell($"B{cnt}").Value = r.Cells["ReferenceNo"].Value.ToString();
//                    //Write data to column C
//                    ws.Cell($"C{cnt}").Value = r.Cells["Date"].Value.ToString();
//                    //Write data to column D
//                    ws.Cell($"D{cnt}").Value = r.Cells["DocType"].Value.ToString();
//                    //Write data to column E
//                    ws.Cell($"E{cnt}").Value = r.Cells["Amount"].Value.ToString().Replace(",", "");
//                    ws.Cell($"E{cnt}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
//                    totAmt = totAmt + Convert.ToDecimal(r.Cells["Amount"].Value.ToString().Replace(",", ""));

//                    cnt += 1;
//                }
//            }
//            #endregion

//            #region STYLE REPORT
//            //Styling
//            lrow = ws.LastRowUsed().RowNumber();
//            ws.Cell($"E{lrow + 1}").Value = totAmt.ToString();
//            ws.Cell($"E{lrow + 1}").Style.Font.Bold = true;
//            ws.Cell($"E{lrow + 1}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";

//            int newRow = lrow + 1;
//            ws.Range($"B{newRow}:E{newRow}").Style.Border.TopBorder = XLBorderStyleValues.Thin;

//            lrow2 = ws.LastRowUsed().RowNumber();

//            IXLRange range = ws.Range(ws.Cell($"B4").Address, ws.Cell($"E{lrow2}").Address);

//            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

//            ws.Columns().AdjustToContents();
//            #endregion

//            workbook.SaveAs(reportFileNme);
//        }


//        MetroSetMessageBox.Show(this, "Reports Generated Successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
//    }
//    catch (Exception ex)
//    {
//        MetroSetMessageBox.Show(this, $"{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
//    }
//}

//public void GenerateNotMatchReportCOPPERWOOD(string vendor)
//{
//    int cnt = 4, lrow = 0, lrow2 = 0;
//    decimal totAmt = 0, totBal = 0;
//    List<string> headerNames = new List<string> { "Transaction Reference", "Document Date", "Document Type", "Amount", "Balance" };
//    try
//    {

//        string reportsDIR = Path.GetDirectoryName(txtStmtFileLoc.Text.Trim()) + $@"\REPORTS\{vendor}";
//        //Create directoy if doesn't exist
//        if (!Directory.Exists(reportsDIR))
//            Directory.CreateDirectory(reportsDIR);

//        string reportFileNme = reportsDIR + $@"\NOTMATCHED_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";

//        string shtnameNOTMATCHED = "NOTMATCHED";
//        string shtnameMATCHED = "MATCHED";
//        //create workbook
//        var isCreated = UtilManager.CreateReportsWorkbook(reportFileNme, shtnameNOTMATCHED, shtnameMATCHED);
//        if (isCreated)
//        {
//            XLWorkbook workbook = new XLWorkbook(reportFileNme);
//            IXLWorksheet ws = workbook.Worksheet(vendor);
//            //Clear worksheet
//            ws.Clear();

//            #region WRITE HEADER
//            //Report Header
//            ws.Cell($"B2").Value = "COPPERWOOD";
//            ws.Cell($"B2").Style.Fill.SetBackgroundColor(XLColor.Blue);
//            // Set the color for the entire cell
//            ws.Cell($"B2").Style.Font.FontColor = XLColor.White;
//            ws.Cell($"B2").Style.Font.Bold = true;
//            ws.Cell($"B2").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
//            //Column headings

//            foreach (var txt in headerNames)
//            {
//                int cntHd = headerNames.IndexOf(txt) + 2;
//                ws.Cell(3, cntHd).Value = txt;
//                ws.Cell(3, cntHd).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
//                ws.Cell(3, cntHd).Style.Fill.SetBackgroundColor(XLColor.Blue);
//                // Set the color for the entire cell
//                ws.Cell(3, cntHd).Style.Font.FontColor = XLColor.White;
//                ws.Cell(3, cntHd).Style.Font.Bold = true;
//            }
//            #endregion

//            #region WRITE DETAILS
//            foreach (DataGridViewRow r in dgSTSMEntries.Rows)
//            {
//                //Look for unmatched statements
//                if (Convert.ToBoolean(r.Cells["chbIsMatch"].Value) == false)
//                {
//                    //Write data to column B
//                    ws.Cell($"B{cnt}").Value = r.Cells["ReferenceNo"].Value.ToString();
//                    //Write data to column C
//                    ws.Cell($"C{cnt}").Value = r.Cells["Date"].Value.ToString();
//                    //Write data to column D
//                    ws.Cell($"D{cnt}").Value = r.Cells["DocType"].Value.ToString();
//                    //Write data to column E
//                    ws.Cell($"E{cnt}").Value = r.Cells["Amount"].Value.ToString().Replace(",", "");
//                    ws.Cell($"E{cnt}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
//                    totAmt = totAmt + Convert.ToDecimal(r.Cells["Amount"].Value.ToString().Replace(",", ""));
//                    //Write data to column E
//                    ws.Cell($"F{cnt}").Value = r.Cells["Balance"].Value.ToString().Replace(",", "");
//                    ws.Cell($"F{cnt}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
//                    totBal = totBal + Convert.ToDecimal(r.Cells["Balance"].Value.ToString().Replace(",", ""));

//                    cnt += 1;
//                }
//            }
//            #endregion

//            #region STYLE REPORT
//            //Styling
//            lrow = ws.LastRowUsed().RowNumber();
//            ws.Cell($"E{lrow + 1}").Value = totAmt.ToString();
//            ws.Cell($"E{lrow + 1}").Style.Font.Bold = true;
//            ws.Cell($"E{lrow + 1}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
//            ws.Cell($"F{lrow + 1}").Value = totBal.ToString();
//            ws.Cell($"F{lrow + 1}").Style.Font.Bold = true;
//            ws.Cell($"F{lrow + 1}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";

//            int newRow = lrow + 1;
//            ws.Range($"B{newRow}:F{newRow}").Style.Border.TopBorder = XLBorderStyleValues.Thin;

//            lrow2 = ws.LastRowUsed().RowNumber();

//            IXLRange range = ws.Range(ws.Cell($"B4").Address, ws.Cell($"F{lrow2}").Address);

//            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

//            ws.Columns().AdjustToContents();
//            #endregion

//            workbook.SaveAs(reportFileNme);
//        }


//        MetroSetMessageBox.Show(this, "Reports Generated Successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
//    }
//    catch (Exception ex)
//    {
//        MetroSetMessageBox.Show(this, $"{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
//    }
//}

//private Task ValidateSTMTEntriesDERRIMON(IProgress<ProgressReport> progress, List<SAPMaster> SAPMasterData)
//{
//    int cntMatch = 0, cntNotMatch = 0, cnt = 1, totalProcess = dgSTSMEntries.Rows.Count;
//    var progressReport = new ProgressReport();

//    dgSTSMEntries.ClearSelection();

//    try
//    {
//        return Task.Run(() =>
//        {
//            foreach (DataGridViewRow row in dgSTSMEntries.Rows)
//            {
//                progressReport.PercentageComplete = cnt++ * 100 / totalProcess;

//                string docNo = row.Cells["ReferenceNo"].Value.ToString();
//                decimal amount = Convert.ToDecimal(row.Cells["Amount"].Value.ToString());
//                DateTime docDate = Convert.ToDateTime(row.Cells["Date"].Value.ToString());

//                if (SAPMasterData.Any(s => s.Reference == docNo && s.Amount == amount.ToString() && s.DocumentDate.Value.Date == docDate.Date))
//                {
//                    row.Cells["chbIsMatch"].Value = CheckState.Checked;
//                    row.Cells["chbIsMatch"].Style.BackColor = System.Drawing.Color.LightGreen;
//                    cntMatch += 1;
//                }
//                else
//                {
//                    row.Cells["chbIsMatch"].Value = CheckState.Unchecked;
//                    row.Cells["chbIsMatch"].Style.BackColor = Color.MistyRose;
//                    cntNotMatch += 1;
//                }

//                progressReport.TotalRecords = dgSTSMEntries.Rows.Count;
//                progressReport.TotalMatched = cntMatch;
//                progressReport.TotalNotMatched = cntNotMatch;

//                progress.Report(progressReport);

//            }



//        });
//    }
//    catch (ThreadInterruptedException e)
//    {
//        MetroSetMessageBox.Show(this, $"{e.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
//        //MessageBox.Show(e.Message);
//        return null;
//    }
//}

//public void GenerateNotMatchReportDERRIMON(string vendor)
//{
//    int cnt = 4, lrow = 0, lrow2 = 0;
//    decimal totAmt = 0;
//    List<string> headerNames = new List<string> { "Transaction Reference", "Document Date", "Document Type", "Amount" };
//    try
//    {

//        string reportsDIR = Path.GetDirectoryName(txtStmtFileLoc.Text.Trim()) + $@"\REPORTS\{vendor}";
//        //Create directoy if doesn't exist
//        if (!Directory.Exists(reportsDIR))
//            Directory.CreateDirectory(reportsDIR);

//        string reportFileNme = reportsDIR + $@"\NOTMATCHED_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";

//        string shtnameNOTMATCHED = "NOTMATCHED";
//        string shtnameMATCHED = "MATCHED";
//        //create workbook
//        var isCreated = UtilManager.CreateReportsWorkbook(reportFileNme, shtnameNOTMATCHED, shtnameMATCHED);
//        if (isCreated)
//        {
//            XLWorkbook workbook = new XLWorkbook(reportFileNme);
//            IXLWorksheet ws = workbook.Worksheet(vendor);
//            //Clear worksheet
//            ws.Clear();

//            #region WRITE HEADER
//            //Report Header
//            ws.Cell($"B2").Value = "DERRIMON TRADING LTD";
//            ws.Cell($"B2").Style.Fill.SetBackgroundColor(XLColor.Blue);
//            // Set the color for the entire cell
//            ws.Cell($"B2").Style.Font.FontColor = XLColor.White;
//            ws.Cell($"B2").Style.Font.Bold = true;
//            ws.Cell($"B2").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
//            //Column headings

//            foreach (var txt in headerNames)
//            {
//                int cntHd = headerNames.IndexOf(txt) + 2;
//                ws.Cell(3, cntHd).Value = txt;
//                ws.Cell(3, cntHd).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
//                ws.Cell(3, cntHd).Style.Fill.SetBackgroundColor(XLColor.Blue);
//                // Set the color for the entire cell
//                ws.Cell(3, cntHd).Style.Font.FontColor = XLColor.White;
//                ws.Cell(3, cntHd).Style.Font.Bold = true;
//            }
//            #endregion

//            #region WRITE DETAILS
//            foreach (DataGridViewRow r in dgSTSMEntries.Rows)
//            {
//                //Look for unmatched statements
//                if (Convert.ToBoolean(r.Cells["chbIsMatch"].Value) == false)
//                {
//                    //Write data to column B
//                    ws.Cell($"B{cnt}").Value = r.Cells["ReferenceNo"].Value.ToString();
//                    //Write data to column C
//                    ws.Cell($"C{cnt}").Value = r.Cells["Date"].Value.ToString();
//                    //Write data to column D
//                    ws.Cell($"D{cnt}").Value = r.Cells["DocType"].Value.ToString();
//                    //Write data to column E
//                    ws.Cell($"E{cnt}").Value = r.Cells["Amount"].Value.ToString().Replace(",", "");
//                    ws.Cell($"E{cnt}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";
//                    totAmt = totAmt + Convert.ToDecimal(r.Cells["Amount"].Value.ToString().Replace(",", ""));

//                    cnt += 1;
//                }
//            }
//            #endregion

//            #region STYLE REPORT
//            //Styling
//            lrow = ws.LastRowUsed().RowNumber();
//            ws.Cell($"E{lrow + 1}").Value = totAmt.ToString();
//            ws.Cell($"E{lrow + 1}").Style.Font.Bold = true;
//            ws.Cell($"E{lrow + 1}").Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00)";

//            int newRow = lrow + 1;
//            ws.Range($"B{newRow}:E{newRow}").Style.Border.TopBorder = XLBorderStyleValues.Thin;

//            lrow2 = ws.LastRowUsed().RowNumber();

//            IXLRange range = ws.Range(ws.Cell($"B4").Address, ws.Cell($"E{lrow2}").Address);

//            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

//            ws.Columns().AdjustToContents();
//            #endregion

//            workbook.SaveAs(reportFileNme);
//        }


//        MetroSetMessageBox.Show(this, "Reports Generated Successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
//    }
//    catch (Exception ex)
//    {
//        MetroSetMessageBox.Show(this, $"{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
//    }
//}