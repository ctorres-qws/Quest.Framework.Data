using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Quest.QuestTools.Presentation.Models
{
    public class JobColorChangeViewModel
    {
        public string Job { get; set; }
        public string ShippingLabelColor { get; set; }
        public bool IsModified { get; set; }
    }
}