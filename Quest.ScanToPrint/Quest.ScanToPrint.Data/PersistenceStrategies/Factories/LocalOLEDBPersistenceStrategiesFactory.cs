using Quest.Framework.Persistance;
using Quest.ScanToPrint.Data.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quest.ScanToPrint.Data
{
    public class LocalOLEDBPersistenceStrategiesFactory : LocalPersistenceStrategiesFactory
    {
        public string DbConnectionString { get; set; }
        public LocalOLEDBPersistenceStrategiesFactory(string dbConnectionString)
        {
            this.DbConnectionString = dbConnectionString;

            this.BarcodesPersistenceStrategy = new OLEDBBarcodesPersistenceStrategy(dbConnectionString);
            this.JobShippingLabelColorPersistenceStrategy = new OLEDBJobShippingLabelColorPersistenceStrategy(dbConnectionString);
            this.LogPersistenceStrategy = new OLEDBLogPersistenceStrategy(dbConnectionString);
            this.RegisteredTagPersistenceStrategy = new OLEDBRegisteredTagPersistenceStrategy(dbConnectionString);
            this.PrinterPersistenceStrategy = new OLEDBPrinterPersistenceStrategy(dbConnectionString);
        }
        protected override IPersistenceStrategy<Barcodes> OnGetBarcodesPersistanceStrategy()
        {
            return this.BarcodesPersistenceStrategy;
        }
        protected override IPersistenceStrategy<JobShippingLabelColor> OnGetJobShippingLabelColorPersistenceStrategy()
        {
            return this.JobShippingLabelColorPersistenceStrategy;
        }
        protected override IPersistenceStrategy<Log> OnGetLogPersistenceStrategy()
        {
            return this.LogPersistenceStrategy;
        }
        protected override IPersistenceStrategy<RegisteredTag> OnGetRegisteredTagPersistenceStrategy()
        {
            return this.RegisteredTagPersistenceStrategy;
        }
        protected override IPersistenceStrategy<Printer> OnGetPrinterPersistenceStrategy()
        {
            return this.PrinterPersistenceStrategy;
        }
    }
}
