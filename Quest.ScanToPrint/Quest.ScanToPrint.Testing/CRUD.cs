using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Quest.ScanToPrint.Data;
using Quest.ScanToPrint.Business;
using Quest.ScanToPrint.Data.Entities;
using System.Collections.Generic;
using System.Printing;
using Quest.ScantToPrint.Presentation;
using System.Management;
using System.Linq;

namespace Quest.ScanToPrint.Testing
{
    [TestClass]
    public class CRUD
    {
        [TestMethod]
        public void CRUDTest()
        {
            FacadeController facadeController = new FacadeController(
                new LocalOLEDBPersistenceStrategiesFactory(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\ScanToPrint\ScanToPrint.mdb;Persist Security Info=False;"),
                new OnlineOLEDBPersistenceStrategiesFactory(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Z:\Quest_Printers.mdb;Persist Security Info=False;")

                );
            int x1 = facadeController.GetTargetPrinter(1);
            int x2 = facadeController.GetTargetPrinter(2);
            int x3 = facadeController.GetTargetPrinter(3);
            //facadeController.AddColorMatch(new ColorMatch() { Color = "Blue", Job = "AAA" });
        }
        [TestMethod]
        public void UploadLocal()
        {
            FacadeController facadeController = new FacadeController(
                new LocalOLEDBPersistenceStrategiesFactory(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\users\ctorres\OneDrive - Quest Window Systems Inc\Desktop\ScanToPrint.mdb;Persist Security Info=False;"),
                new OnlineOLEDBPersistenceStrategiesFactory(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=V:\Quest_CT.mdb;Persist Security Info=False;")

                );
            //List<RegisteredTag> tags = facadeController.GetRegisteredTags().Where(x => x.Job == "AAA").ToList();

            //foreach(RegisteredTag tag in tags)
            //{
            //    facadeController.AddBarcode(new Barcodes()
            //    {
            //        Barcode = string.Format("{0}{1}{2}", tag.Job, tag.Floor, tag.Tag),
            //        Tag = tag.Tag,
            //        Job = tag.Job,
            //        Line = 2,
            //        ScanDate = DateTime.Now,
            //        SentDatabase = false
            //    });
            //}
            

            facadeController.UploadBarcodes();
        }

        [TestMethod]
        public void TestMail()
        {
            FacadeController facadeController = new FacadeController(
                new LocalOLEDBPersistenceStrategiesFactory(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\ScanToPrint\ScanToPrint.mdb;Persist Security Info=False;"),
                new OnlineOLEDBPersistenceStrategiesFactory(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Z:\Quest.mdb;Persist Security Info=False;")

                );

            facadeController.GetJobOrderEntryData("AAA");
            List<Glazing> xxx = facadeController.GetGlazings();
        }
        [TestMethod]
        public void Printing()
        {
            var server = new LocalPrintServer();

            PrintQueue queue = server.GetPrintQueue(@"\\QWTORPRINT1\Goreway-Glazing1", new string[0] { });
        }
        [TestMethod]
        public void PDF()
        { 
            PDFcreator pDFcreator = new PDFcreator();

            pDFcreator.CreatePDF6x4FromAnotherLine(new BarcodeReading() { Barcode = "AAA", Floor = "13A", Job = "AAA", Tag = "-001", Line = 3, ScanDate = DateTime.Now }, "#7f0000");
        }
        [TestMethod]
        public void PrintersCheking()
        {
            ManagementObjectSearcher searcher = new
      ManagementObjectSearcher("SELECT * FROM   Win32_Printer");

            bool IsReady = false;
            foreach (ManagementObject printer in searcher.Get())
            {
                if (printer["Name"].ToString().ToLower().Contains("goreway-glazing1"))
                {
                    PropertyDataCollection dataCollection = printer.Properties;

                    Dictionary<object, object> dict = new Dictionary<object, object>();

                    foreach(PropertyData propertyData in dataCollection)
                    {
                        dict.Add(propertyData.Name, propertyData.Value);
                    }
                //    if (printer["WorkOffline"].ToString().ToLower().Equals("false"))
                //    {
                //        IsReady = true;
                    }
                //}
            }
        }
    }
}
