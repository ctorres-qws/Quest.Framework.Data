using Quest.Framework.Persistance;
using Quest.ScanToPrint.Data.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quest.ScanToPrint.Data
{
    public class OnlineOLEDBPersistenceStrategiesFactory : OnlinePersistenceStrategiesFactory
    {
        public string DbConnectionString { get; set; }
        public OnlineOLEDBPersistenceStrategiesFactory(string dbConnectionString)
        {
            this.DbConnectionString = dbConnectionString;

            this.GlazingPersistenceStrategy = new OLEDBGlazingPersistenceStrategy(dbConnectionString);
            this.JobPersistenceStrategy = new OLEDBJobPersistenceStrategy(dbConnectionString);
            this.JobOrderEntryDataPersistenceStrategy = new OLEDBJobOrderEntryDataPersistenceStrategy(dbConnectionString);
            this.PrinterPersistenceStrategy = new OLEDBPrinterPersistenceStrategy(dbConnectionString);
        }
        protected override IPersistenceStrategy<Glazing> OnGetGlazingPersistenceStrategy()
        {
            return this.GlazingPersistenceStrategy;
        }
        protected override IPersistenceStrategy<Job> OnGetJobPersistenceStrategy()
        {
            return this.JobPersistenceStrategy;
        }
        protected override IPersistenceStrategy<JobOrderEntryData> OnGetJobOrderEntryDataPersistenceStrategy()
        {
            return this.JobOrderEntryDataPersistenceStrategy;
        }
        protected override IPersistenceStrategy<Printer> OnGetPrinterPersistenceStrategy()
        {
            return this.PrinterPersistenceStrategy;
        }
    }
}
