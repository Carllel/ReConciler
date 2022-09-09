using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReConciler.Model
{
    public class vendorStatement
    {
        public string ReferenceNo { get; set; }
        public string DocumentDate { get; set; }
        public string DocumentType { get; set; }
        public string Amount { get; set; }
        public string Balance { get; set; }
    }

    public class validateVendorStatement
    {
        public string RefNo { get; set; }
        public string DocDate { get; set; }
        public string Amt { get; set; }
        public string Bal { get; set; }
    }
}
