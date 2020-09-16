using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quest.ScanToPrint.Data.Entities
{
    public class Glazing
    {
        public string Barcode { get; set; }
        public string Job { get; set; }
        public string Floor { get; set; }
        public string Tag { get; set; }
        public string Dept { get; set; }
        public string Employee { get; set; }
        public int Openings { get; set; }
        public string FirstComplete { get; set; }
        public int Joints { get; set; }
        public DateTime DateTime { get; set; }
        public int Day { get; set; }
        public int Month { get; set; }
        public int Year { get; set; }
        public TimeSpan Time { get; set; }
        public int Week { get; set; }
        public int ONumber { get; set; }
        public int ScanCount { get; set; }
        public string O1 { get; set; }
        public string O2 { get; set; }
        public string O3 { get; set; }
        public string O4 { get; set; }
        public string O5 { get; set; }
        public string O6 { get; set; }
        public string O7 { get; set; }
        public string O8 { get; set; }
        public int Count { get; set; }
    }
}
