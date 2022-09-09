using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReConciler.Model
{
    public class ProgressReport
    {
        public int PercentageComplete { get; set; }
        public int TotalRecords { get; set; }
        public int TotalMatched { get; set; }
        public int TotalNotMatched { get; set; }
    }
}
