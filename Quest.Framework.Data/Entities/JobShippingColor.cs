using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quest.Framework.Data.Entities
{
    public class JobShippingColor
    {
        public string Job { get; set; }
        public string Parent { get; set; }
        public string ShippingLabelColor { get; set; }
        public string ColorName { get; set; }
    }
}
