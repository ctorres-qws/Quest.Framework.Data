using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quest.ScanToPrint.Data.Entities
{
    public class Barcodes
    {
        public int ID { get; set; }
        public string Job { get; set; }
        public string Barcode { get; set; }
        public string Tag { get; set; }
        public DateTime ScanDate { get; set; }
        public int Line { get; set; }
        public bool SentPrint { get; set; }
        public bool SentDatabase { get; set; }
    }
}
