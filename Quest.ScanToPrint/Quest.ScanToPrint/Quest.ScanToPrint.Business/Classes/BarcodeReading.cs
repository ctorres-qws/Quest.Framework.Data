using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quest.ScanToPrint.Business
{
    public class BarcodeReading
    {
        public string Barcode { get ; set; }
        public int Line { get; set; }
        public string Job { get; set; }
        public string Floor { get; set; }
        public string Tag { get; set; }
        public DateTime ScanDate { get; set; }
    }
}
