using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Quest.ScanToPrint.Data;
using Quest.ScanToPrint.Business;
using Quest.ScanToPrint.Data.Entities;
using System.Collections.Generic;
using System.Printing;
using Quest.ScantToPrint.Presentation;
using System.Management;

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
                new OnlineOLEDBPersistenceStrategiesFactory(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Z:\Quest_ScanToPrintDevelopment.mdb;Persist Security Info=False;")

                );
            List<Glazing> glazings = facadeController.GetGlazings();
            //facadeController.AddColorMatch(new ColorMatch() { Color = "Blue", Job = "AAA" });
        }
        [TestMethod]
        public void UploadLocal()
        {
            FacadeController facadeController = new FacadeController(
                new LocalOLEDBPersistenceStrategiesFactory(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\ScanToPrint\ScanToPrint.mdb;Persist Security Info=False;"),
                new OnlineOLEDBPersistenceStrategiesFactory(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Z:\Quest.mdb;Persist Security Info=False;")

                );


            facadeController.UploadBarcodes();
        }

        [TestMethod]
        public void TestMail()
        {
            FacadeController facadeController = new FacadeController(
                new LocalOLEDBPersistenceStrategiesFactory(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\ScanToPrint\ScanToPrint.mdb;Persist Security Info=False;"),
                new OnlineOLEDBPersistenceStrategiesFactory(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Z:\Quest.mdb;Persist Security Info=False;")

                );


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

            pDFcreator.CreatePDF6x4(new BarcodeReading() { Barcode = "AAA", Floor = "2A", Job = "AAA", Tag = "-134" }, "#7f0000");
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
