using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Quest.Framework.Data.Entities;
using Quest.Framework.Data.Persistence.OLEDB;
using Quest.QuestTools.Business;

namespace Quest.QuestTools.Testing
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            FacadeController controller = new FacadeController(new OLEDBPersistenceStrategiesFactory(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\users\ctorres\OneDrive - Quest Window Systems Inc\Desktop\Quest_ScanToPrintDevelopment.mdb;Persist Security Info=False;"));

            //List<Job> jobs = controller.GetJobs();
            //List<ShippingColor> sc = controller.GetShippingColors();
            List<JobShippingColor> jsc = controller.GetJobShippingColors();
        }
    }
}
