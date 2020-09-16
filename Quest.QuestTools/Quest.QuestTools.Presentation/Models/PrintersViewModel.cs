using Quest.Framework.Data.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Quest.QuestTools.Presentation.Models
{
    public class PrintersViewModel
    {
        public List<Printer> Printers { get; set; }
        public User User { get; set; }
    }
}