
using LinqToExcel;
using ReConciler.Model;
using ReConciler.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReConciler.Data
{
    public static class DataAccess
    {
        public static bool IsDouble(string text)
        {
            Double num = 0;
            bool isDouble = false;

            // Check for empty string.
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            isDouble = Double.TryParse(text, out num);

            return isDouble;
        }
        #region LOAD VENDOR STATEMENTS FROM TEMPLATE

        public static List<vendorSTMT> GetSTMTEntriesMassy(string vendor, string fpath)
        {
            try
            {
                List<vendorSTMT> sTMTs = new List<vendorSTMT>();

                var excel = new ExcelQueryFactory(fpath);
                excel.ReadOnly = true;
                var lstEntries = from c in excel.WorksheetNoHeader(vendor)
                                 where c[3] == "I" || c[3] == "C" //where 3 is the column index of the worksheet
                                 select c;

                foreach (var a in lstEntries)
                {
                    string amt = a[6].Value.ToString();
                    amt = amt.Length == 0 ? "0.00" : amt.Trim();

                    sTMTs.Add(new vendorSTMT() { ReferenceNo = a[1], DocumentDate = a[0], DocumentType = a[3], Amount = amt });
                }

                return sTMTs;

            }
            catch (Exception ex)
            {

                return null;
            }
        }
        //public static List<massySTMT> GetSTMTEntriesMassyV2(string vendor, string fpath)
        //{

        //    try
        //    {
        //        List<massySTMT> massySTMTs = new List<massySTMT>();

        //        var excel = new ExcelQueryFactory(fpath);
        //        excel.ReadOnly = true;
        //        var lstEntries = from c in excel.WorksheetNoHeader(vendor)
        //                         where c[3] == "I" || c[3] == "C" //where 3 is the column index of the worksheet
        //                         select c;

        //        foreach (var a in lstEntries)
        //        {
        //            massySTMTs.Add(new massySTMT() { TransDate = a[0], TransNo = a[1], CustomerPO = a[2],
        //                DocType = a[3], DueDate = a[4], OriginalAmount = a[5], OutstandingAmount = a[6] });
        //        }

        //        return massySTMTs;

        //    }
        //    catch (Exception ex)
        //    {

        //        return null;
        //    }
        //}
        public static List<wisyncoSTMT> GetSTMTEntriesWisynco(string vendor,string fpath)
        {
            //var jsonFIleFUllPath = AppDomain.CurrentDomain.BaseDirectory + @"\DocumentTypes.json";
            //var documents = UtilManager.GetDocumentTypes(jsonFIleFUllPath).ToList();

            try
            {
                List<wisyncoSTMT> wisyncoSTMTs = new List<wisyncoSTMT>();

                var excel = new ExcelQueryFactory(fpath);
                excel.ReadOnly = true;
                var lstEntries = from c in excel.WorksheetNoHeader(vendor)
                                 where c[2] == "INV" || c[2] == "ADJ" || c[2] == "PAY" || c[2] == "CRD"//where 2 is the column index of the worksheet
                                 select c;

                foreach (var a in lstEntries)
                {
                    string drAmt = a[3].Value.ToString().Length == 0 ? "0.00" : a[3].Value.ToString();
                    string crAmt = a[4].Value.ToString().Length == 0 ? "0.00" : a[4].Value.ToString();
                    string bal = a[5].Value.ToString().Length == 0 ? "0.00" : a[5].Value.ToString();
                    //Remove minus sign from the end of the amount and place it at the start
                    if (drAmt.Contains("-"))
                        drAmt = drAmt.Substring(drAmt.Length - 1) + drAmt.Remove(drAmt.Length - 1);
                    if (crAmt.Contains("-"))
                        crAmt = crAmt.Substring(crAmt.Length - 1) + crAmt.Remove(crAmt.Length - 1);
                    if (bal.Contains("-"))
                        bal = bal.Substring(bal.Length - 1) + bal.Remove(bal.Length - 1);
                    //string docName = documents.Where(d => d.docId == a[2]).Select(t => t.docName).SingleOrDefault();
                    wisyncoSTMTs.Add(new wisyncoSTMT() { ReferenceNo = a[0], Date = a[1], DocType = a[2], Debit = drAmt, Credit = crAmt, Balance = bal });
                }

                return wisyncoSTMTs;

            }
            catch (Exception ex)
            {

                return null;
            }
        }
        public static List<caribProducerSTMT> GetSTMTEntriesCaribProducers(string vendor,string fpath)
        {
            try
            {
                List<caribProducerSTMT> sTMTs = new List<caribProducerSTMT>();

                var excel = new ExcelQueryFactory(fpath);
                excel.ReadOnly = true;
                var lstEntries = from c in excel.WorksheetNoHeader(vendor)
                                 where c[2] == "SLS" || c[2] == "SCH" || c[2] == "DR"
                                    || c[2] == "CR" || c[2] == "RTN" || c[2] == "PMT"//where 2 is the column index of the worksheet
                                 select c;

                foreach (var a in lstEntries)
                {
                    string amt = a[4].Value.ToString().Replace("J$", "");
                    amt = amt.Length == 0 ? "0.00" : amt.Trim();
                    string bal = a[5].Value.ToString().Trim().Replace("J$", "");
                    bal = bal.Length == 0 ? "0.00" : bal.Trim();

                    sTMTs.Add(new caribProducerSTMT() { ReferenceNo = a[0], Date = a[1], DocType = a[2], Amount = amt, Balance = bal });
                }

                return sTMTs;

            }
            catch (Exception ex)
            {

                return null;
            }
        }
        public static List<vendorSTMT> GetSTMTEntriesConsolBakeries(string vendor,string fpath)
        {
            try
            {
                List<vendorSTMT> sTMTs = new List<vendorSTMT>();

                var excel = new ExcelQueryFactory(fpath);
                excel.ReadOnly = true;
                var lstEntries = from c in excel.WorksheetNoHeader(vendor)
                                 where c[3] == "DB" || c[3] == "PY" || c[3] == "UC"
                                    || c[3] == "CR" || c[3] == "ED" || c[3] == "RF"//where 3 is the column index of the worksheet
                                    || c[3] == "IT" || c[3] == "PI" || c[3] == "AD" || c[3] == "IN"
                                 select c;

                foreach (var a in lstEntries)
                {
                    string amt = a[6].Value.ToString();
                    amt = amt.Length == 0 ? "0.00" : amt.Trim();

                    sTMTs.Add(new vendorSTMT() { ReferenceNo = a[0], DocumentDate = a[2], DocumentType = a[3], Amount = amt });
                }

                return sTMTs;

            }
            catch (Exception ex)
            {

                return null;
            }
        }
        public static List<vendorSTMT> GetSTMTEntriesContinentalBaking(string vendor,string fpath)
        {
            try
            {
                List<vendorSTMT> sTMTs = new List<vendorSTMT>();

                var excel = new ExcelQueryFactory(fpath);
                excel.ReadOnly = true;
                var lstEntries = from c in excel.WorksheetNoHeader(vendor)
                                 where c[2] == "DB" || c[2] == "PY" || c[2] == "UC"
                                    || c[2] == "CR" || c[2] == "ED" || c[2] == "RF"//where 3 is the column index of the worksheet
                                    || c[2] == "IT" || c[2] == "PI" || c[2] == "AD" || c[2] == "IN"
                                 select c;

                foreach (var a in lstEntries)
                {
                    string amt = a[5].Value.ToString();
                    amt = amt.Length == 0 ? "0.00" : amt.Trim();

                    sTMTs.Add(new vendorSTMT() { ReferenceNo = a[0], DocumentDate = a[1], DocumentType = a[2], Amount = amt });
                }

                return sTMTs;

            }
            catch (Exception ex)
            {

                return null;
            }
        }
        public static List<rainforestSTMT> GetSTMTEntriesRainforest(string vendor,string fpath)
        {
            try
            {
                List<rainforestSTMT> sTMTs = new List<rainforestSTMT>();

                var excel = new ExcelQueryFactory(fpath);
                excel.ReadOnly = true;
                var lstEntries = from c in excel.WorksheetNoHeader(vendor)
                                 where c[7] != "" && c[0] != ""
                                 select c;

                foreach (var a in lstEntries)
                {
                    string invamt = a[5].Value.ToString();
                    invamt = invamt.Length == 0 ? "0.00" : invamt.Trim();
                    string pmtamt = a[6].Value.ToString();
                    pmtamt = pmtamt.Length == 0 ? "0.00" : pmtamt.Trim();
                    string bal = a[7].Value.ToString();
                    bal = bal.Length == 0 ? "0.00" : bal.Trim();

                    //Remove minus sign from the end of the amount and place it at the start
                    if (invamt.Contains("-"))
                        invamt = invamt.Substring(invamt.Length - 1) + invamt.Remove(invamt.Length - 1);
                    if (pmtamt.Contains("-"))
                        pmtamt = pmtamt.Substring(pmtamt.Length - 1) + pmtamt.Remove(pmtamt.Length - 1);
                    if (bal.Contains("-"))
                        bal = bal.Substring(bal.Length - 1) + bal.Remove(bal.Length - 1);

                    sTMTs.Add(new rainforestSTMT() { ReferenceNo = a[0], Date = a[1], InvoiceAmount = invamt, Payments = pmtamt, Balance = bal });
                }

                return sTMTs;

            }
            catch (Exception ex)
            {

                return null;
            }
        }
        public static List<vendorSTMT> GetSTMTEntriesWorldBrands(string vendor,string fpath)
        {
            try
            {
                List<vendorSTMT> sTMTs = new List<vendorSTMT>();

                var excel = new ExcelQueryFactory(fpath);
                excel.ReadOnly = true;
                var lstEntries = from c in excel.WorksheetNoHeader(vendor)
                                 where c[5] != "" && c[2] != ""
                                 select c;

                foreach (var a in lstEntries)
                {
                    string amt = a[5].Value.ToString();
                    amt = amt.Length == 0 ? "0.00" : amt.Trim();
                    if (IsDouble(amt))
                    {
                        sTMTs.Add(new vendorSTMT() { ReferenceNo = a[2], DocumentDate = a[4], DocumentType = a[3], Amount = amt });
                    }
                }

                return sTMTs;

            }
            catch (Exception ex)
            {

                return null;
            }
        }
        public static List<vendorSTMT> GetSTMTEntriesConsumerBrands(string vendor,string fpath)
        {
            try
            {
                List<vendorSTMT> sTMTs = new List<vendorSTMT>();

                var excel = new ExcelQueryFactory(fpath);
                excel.ReadOnly = true;
                var lstEntries = from c in excel.WorksheetNoHeader(vendor)
                                 where c[1] != ""
                                 select c;

                foreach (var a in lstEntries)
                {
                    string amt = a[4].Value.ToString();
                    amt = amt.Length == 0 ? "0.00" : amt.Trim().Replace("(", "-").Replace(")", "");
                    if (!amt.Contains("Amt in loc.cur.") || !a[1].Value.ToString().Contains("Document"))
                    {
                        sTMTs.Add(new vendorSTMT() { ReferenceNo = a[1], DocumentDate = a[3], DocumentType = a[2], Amount = amt });
                    }
                    //sTMTs.Add(new vendorSTMT() { ReferenceNo = a[1], Date = a[3], DocType = a[2], Amount = amt });
                }

                return sTMTs;

            }
            catch (Exception ex)
            {

                return null;
            }
        }
        public static List<vendorStatement> GetSTMTEntriesFaceyPharmacy(string vendor,string fpath)
        {
            try
            {
                List<vendorStatement> sTMTs = new List<vendorStatement>();

                var excel = new ExcelQueryFactory(fpath);
                excel.ReadOnly = true;
                var lstEntries = from c in excel.WorksheetNoHeader(vendor)
                                 where c[1] != "" && c[1] != "Typ"
                                 select c;

                foreach (var a in lstEntries)
                {
                    string amt = a[5].Value.ToString();
                    string bal = a[6].Value.ToString();
                    amt = amt.Length == 0 ? "0.00" : amt.Trim();
                    bal = bal.Length == 0 ? "0.00" : bal.Trim();

                    sTMTs.Add(new vendorStatement() { ReferenceNo = a[0], DocumentDate = a[2], DocumentType = a[1], Amount = amt, Balance = bal });
                }

                return sTMTs;

            }
            catch (Exception ex)
            {

                return null;
            }
        }
        public static List<vendorStatement> GetSTMTEntriesCarimed(string vendor, string fpath)
        {
            try
            {
                List<vendorStatement> sTMTs = new List<vendorStatement>();

                var excel = new ExcelQueryFactory(fpath);
                excel.ReadOnly = true;
                var lstEntries = from c in excel.WorksheetNoHeader(vendor)
                                 where c[5] != "" && c[5] != "Original Trx Amount"
                                 select c;

                foreach (var a in lstEntries)
                {
                    string amt = a[5].Value.ToString();
                    string bal = a[6].Value.ToString();
                    amt = amt.Length == 0 ? "0.00" : amt.Trim();
                    bal = bal.Length == 0 ? "0.00" : bal.Trim();
                    //var docdate = DateTime.FromOADate(Convert.ToDouble(a[4].Value)); docdate.ToString("dd/MM/yyyy")
                    sTMTs.Add(new vendorStatement() { ReferenceNo = a[0], DocumentDate = a[4], DocumentType = a[3], Amount = amt, Balance = bal });
                }

                return sTMTs;

            }
            catch (Exception ex)
            {

                return null;
            }
        }
        public static List<vendorStatement> GetSTMTEntriesFacey(string vendor, string fpath)
        {
            try
            {
                List<vendorStatement> sTMTs = new List<vendorStatement>();

                var excel = new ExcelQueryFactory(fpath);
                //excel.ReadOnly = true;
                var lstEntries = from c in excel.WorksheetNoHeader(vendor)
                                 where c[1] == "Ovp" || c[1] == "Inv" || c[1] == "Pmt" || c[1] == "Crd" || c[1] == "Crd"
                                 //where c[5] != "" && c[5] != "Transaction Amount"
                                 select c;

                foreach (var a in lstEntries)
                {
                    string amt = a[5].Value.ToString().Replace("$", "");
                    string bal = a[6].Value.ToString().Replace("$", "");
                    amt = amt.Length == 0 ? "0.00" : amt.Trim();
                    bal = bal.Length == 0 ? "0.00" : bal.Trim();
                    //var docdate = DateTime.FromOADate(Convert.ToDouble(a[4].Value));
                    sTMTs.Add(new vendorStatement() { ReferenceNo = a[0], DocumentDate = a[2], DocumentType = a[1], Amount = amt, Balance = bal });
                }

                return sTMTs;

            }
            catch (Exception ex)
            {

                return null;
            }
        }
        public static List<vendorSTMT> GetSTMTEntriesDerrimon(string vendor, string fpath)
        {
            try
            {
                List<vendorSTMT> sTMTs = new List<vendorSTMT>();

                var excel = new ExcelQueryFactory(fpath);
                excel.ReadOnly = true;
                var lstEntries = from c in excel.WorksheetNoHeader(vendor)
                                 where c[6] != "" && c[2] != "" && c[6] != "Remaining Amt. ($)"
                                 select c;

                foreach (var a in lstEntries)
                {
                    string amt = a[6].Value.ToString();
                    amt = amt.Length == 0 ? "0.00" : amt.Trim();
                    //var docdate = DateTime.FromOADate(Convert.ToDouble(a[0].Value)); docdate.ToString("dd/MM/yyyy")
                    sTMTs.Add(new vendorSTMT() { ReferenceNo = a[2], DocumentDate = a[0], DocumentType = a[1], Amount = amt });
                }

                return sTMTs;

            }
            catch (Exception ex)
            {

                return null;
            }
        }
        public static List<vendorSTMTv2> GetSTMTEntriesBestDressed(string vendor, string fpath)
        {
            try
            {
                ///Check if exel file is open
                List<vendorSTMTv2> sTMTs = new List<vendorSTMTv2>();

                var excel = new ExcelQueryFactory(fpath);
                excel.ReadOnly = false;
                var lstEntries = from c in excel.WorksheetNoHeader(vendor)
                                 where c[5] == "JMD"  //Filter by currency column 
                                 select c;

                foreach (var a in lstEntries)
                {
                    string drAmt = a[6].Value.ToString().Length == 0 ? "0.00" : a[6].Value.ToString();
                    string crAmt = a[7].Value.ToString().Length == 0 ? "0.00" : a[7].Value.ToString();
                    string bal = a[8].Value.ToString().Length == 0 ? "0.00" : a[8].Value.ToString();

                    //var docdate = DateTime.FromOADate(Convert.ToDouble(a[0].Value));docdate.ToString("dd/MM/yyyy")
                    sTMTs.Add(new vendorSTMTv2() { ReferenceNo = a[1], Date = a[0], DocType = "", Debit = drAmt, Credit = crAmt, Balance = bal });
                }

                return sTMTs;

            }
            catch (Exception ex)
            {

                return null;
            }
        }
        public static List<vendorStatement> GetSTMTEntriesCopperwood(string vendor, string fpath)
        {
            try
            {
                List<vendorStatement> sTMTs = new List<vendorStatement>();

                var excel = new ExcelQueryFactory(fpath);
                excel.ReadOnly = true;
                var lstEntries = from c in excel.WorksheetNoHeader(vendor)
                                 where c[0] == "IN" || c[0] == "DN" || c[0] == "CN" || c[0] == "CR"//where 0 is the column index of the worksheet
                                 select c;

                foreach (var a in lstEntries)
                {
                    string amt = a[5].Value.ToString().Length == 0 ? "0.00" : a[5].Value.ToString();
                    string bal = a[6].Value.ToString().Length == 0 ? "0.00" : a[6].Value.ToString();

                    //var docdate = DateTime.FromOADate(Convert.ToDouble(a[2].Value));docdate.ToString("dd/MM/yyyy")
                    sTMTs.Add(new vendorStatement() { ReferenceNo = a[1], DocumentDate = a[2], DocumentType = a[0], Amount = amt, Balance = bal });
                }

                return sTMTs;

            }
            catch (Exception ex)
            {

                return null;
            }
        }
        public static List<vendorStatement> GetSTMTEntriesTGeddes(string vendor, string fpath)
        {
            try
            {
                List<vendorStatement> sTMTs = new List<vendorStatement>();

                var excel = new ExcelQueryFactory(fpath);
                excel.ReadOnly = true;
                var lstEntries = from c in excel.WorksheetNoHeader(vendor)
                                 where c[1] == "Crd" || c[1] == "Inv" || c[1] == "Pmt" || c[1] == "Ovp"//where 1 is the column index of the worksheet
                                 select c;

                foreach (var a in lstEntries)
                {
                    string amt = a[5].Value.ToString().Length == 0 ? "0.00" : a[5].Value.ToString();
                    string bal = a[6].Value.ToString().Length == 0 ? "0.00" : a[6].Value.ToString();

                    //var docdate = DateTime.FromOADate(Convert.ToDouble(a[2].Value));docdate.ToString("dd/MM/yyyy")
                    sTMTs.Add(new vendorStatement() { ReferenceNo = a[0], DocumentDate = a[2], DocumentType = a[1], Amount = amt, Balance = bal });
                }

                return sTMTs;

            }
            catch (Exception ex)
            {

                return null;
            }
        }
        public static List<vendorSTMTv2> GetSTMTEntriesChasERamson(string vendor, string fpath)
        {
            try
            {
                List<vendorSTMTv2> sTMTs = new List<vendorSTMTv2>();

                var excel = new ExcelQueryFactory(fpath);
                excel.ReadOnly = true;
                var lstEntries = from c in excel.WorksheetNoHeader(vendor)
                                 where c[2] == "CR" || c[2] == "PY" || c[2] == "UC" || c[2] == "IN"//where 1 is the column index of the worksheet
                                 select c;

                foreach (var a in lstEntries)
                {
                    string drAmt = a[4].Value.ToString().Length == 0 ? "0.00" : a[4].Value.ToString();
                    string crAmt = a[5].Value.ToString().Length == 0 ? "0.00" : a[5].Value.ToString();
                    string bal = a[6].Value.ToString().Length == 0 ? "0.00" : a[6].Value.ToString();

                    //var docdate = DateTime.FromOADate(Convert.ToDouble(a[2].Value));docdate.ToString("dd/MM/yyyy")
                    sTMTs.Add(new vendorSTMTv2() { ReferenceNo = a[3], Date = a[0], DocType = a[2], Debit = drAmt, Credit = crAmt, Balance = bal });
                }

                return sTMTs;

            }
            catch (Exception ex)
            {

                return null;
            }
        }
        public static List<vendorSTMTv2> GetSTMTEntriesCeleBrands(string vendor, string fpath)
        {
            try
            {
                List<vendorSTMTv2> sTMTs = new List<vendorSTMTv2>();

                var excel = new ExcelQueryFactory(fpath);
                excel.ReadOnly = true;
                var lstEntries = from c in excel.WorksheetNoHeader(vendor)
                                 where c[2] == "CR" || c[2] == "PY" || c[2] == "UC" || c[2] == "IN"//where 1 is the column index of the worksheet
                                 select c;

                foreach (var a in lstEntries)
                {
                    string drAmt = a[4].Value.ToString().Length == 0 ? "0.00" : a[4].Value.ToString();
                    string crAmt = a[5].Value.ToString().Length == 0 ? "0.00" : a[5].Value.ToString();
                    string bal = a[6].Value.ToString().Length == 0 ? "0.00" : a[6].Value.ToString();

                    //var docdate = DateTime.FromOADate(Convert.ToDouble(a[2].Value));docdate.ToString("dd/MM/yyyy")
                    sTMTs.Add(new vendorSTMTv2() { ReferenceNo = a[3], Date = a[0], DocType = a[2], Debit = drAmt, Credit = crAmt, Balance = bal });
                }

                return sTMTs;

            }
            catch (Exception ex)
            {

                return null;
            }
        }

        public static List<product> GetProductList(string fpath)
        {
            List<product> products = new List<product>();
            var excel = new ExcelQueryFactory(fpath);
            excel.ReadOnly = true;

            var lstEntries = (from p in excel.WorksheetNoHeader("Sheet1")
                              where p[1] != "" && p[1] != "Description"
                              select p).ToList();


            foreach(var entry in lstEntries)
            {
                string newpv = entry[6] == "-" ? "0.00" : entry[6];
                string newbv = entry[7] == "-" ? "0.00" : entry[7];
                string newibo = entry[8] == "-" ? "0.00" : entry[8];
                string newrtval = entry[9] == "-" ? "0.00" : entry[9];
                products.Add(new product() { productId = entry[0], description = entry[1], pv = newpv, bv = newbv, ibovalue = newibo, retailvalue = newrtval });
            }

            return products;

        }

        #endregion


        #region GET SAP MASTER DATA
        public static List<SAPMaster> GetSAPMasterData(string sheetName,string fpath)
        {
            List<SAPMaster> sapdata = new List<SAPMaster>();
            DateTime? docdate = null;
            DateTime? postdate = null;
            try
            {
                var excel = new ExcelQueryFactory(fpath);
                excel.ReadOnly = true;
                var lstEntries = (from c in excel.Worksheet(sheetName)
                                  where c["Reference"] != ""
                                  select c).ToList();

                foreach (var a in lstEntries)
                {
                    
                    if (a[2].Value.ToString().Contains("."))//Handle date format dd.MM.yyyy
                    {
                        var splitDocDte = a[2].Value.ToString().Split('.');
                        var newDocDte = splitDocDte[1] + @"/" + splitDocDte[0] + @"/" + splitDocDte[2];
                        docdate = Convert.ToDateTime(newDocDte); 
                    }
                    else
                    {
                        docdate = Convert.ToDateTime(a[2].Value.ToString());
                        
                    }

                    if (a[6].Value.ToString().Contains("."))//Handle date format dd.MM.yyyy
                    {
                        var splitPostDte = a[6].Value.ToString().Split('.');
                        var newPosDte = splitPostDte[1] + @"/" + splitPostDte[0] + @"/" + splitPostDte[2];
                        postdate = Convert.ToDateTime(newPosDte);
                    }
                    else
                    {
                        postdate = Convert.ToDateTime(a[6].Value.ToString());

                    }

                    

                    
                    //string crAmt = a[5].Value.ToString().Length == 0 ? "0.00" : a[5].Value.ToString();
                    //string bal = a[6].Value.ToString().Length == 0 ? "0.00" : a[6].Value.ToString();

                    sapdata.Add(new SAPMaster()
                    {
                        Reference = a[0],
                        DocumentNum = a[1],
                        DocumentDate = docdate,
                        Amount =  a[3],
                        CheckNum = a[4],
                        DocumentHeader = a[5],
                        PostingDate = postdate,
                        Description = a[7],
                        User = a[8]
                    });
                }

                return sapdata;

            }
            catch (Exception ex)
            {
                //MetroSetMessageBox.Show(this, $"{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return sapdata;
            }
        }
        #endregion
    }
}
