using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReConciler.Model
{
    public class vendorSTMTv2
    {
        public string ReferenceNo { get; set; }
        public string Date { get; set; }
        public string DocType { get; set; }
        public string Debit { get; set; }
        public string Credit { get; set; }
        public string Balance { get; set; }
    }

    public class validateVendorSTMTv2
    {
        public string RefNo { get; set; }
        public string DocDate { get; set; }
        public string DocType { get; set; }
        public string DebitTotal { get; set; }
        public string CreditTotal { get; set; }
        public string BalTotal { get; set; }
    }
}
