using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReConciler.Model
{
    public class SAPMaster
    {
        public string Reference { get; set; }
        public string DocumentNum { get; set; }
        public DateTime? DocumentDate { get; set; }
        public string Amount { get; set; }
        public string CheckNum { get; set; }
        public string DocumentHeader { get; set; }
        public DateTime? PostingDate { get; set; }
        public string Description { get; set; }
        public string User { get; set; }
    }
}
