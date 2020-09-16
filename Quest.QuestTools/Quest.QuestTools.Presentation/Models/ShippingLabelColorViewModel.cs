using Quest.Framework.Data.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Quest.QuestTools.Presentation.Models
{
    public class ShippingLabelColorViewModel
    {
        public List<ColorCatalogViewModel> Colors { get; set; }
        public List<JobShippingColor> JobShippingColors { get; set; }
        public List<Job> Jobs { get; set; }
        public User User { get; set; }
    }
}