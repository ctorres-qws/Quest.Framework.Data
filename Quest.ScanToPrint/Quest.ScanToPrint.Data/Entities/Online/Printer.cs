using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quest.ScanToPrint.Data.Entities
{
    public class Printer
    {
        public int ID { get; set; }
        public int GlazingLine { get; set; }
        public bool Active { get; set; }
        public int BackupPrinter { get; set; }
    }
}
