using Quest.Framework.Persistance;
using Quest.ScanToPrint.Data.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quest.ScanToPrint.Data
{
    public abstract class OnlinePersistenceStrategiesFactory : IPersistenceStrategiesFactory
    {
        internal IPersistenceStrategy<Glazing> GlazingPersistenceStrategy { get; set; }
        internal IPersistenceStrategy<Job> JobPersistenceStrategy { get; set; }
        internal IPersistenceStrategy<JobOrderEntryData> JobOrderEntryDataPersistenceStrategy { get; set; }
        internal IPersistenceStrategy<Printer> PrinterPersistenceStrategy { get; set; }
        public IPersistenceStrategy<Glazing> GetGlazingPersistenceStrategy()
        {
            return OnGetGlazingPersistenceStrategy();
        }
        public IPersistenceStrategy<Job> GetJobPersistenceStrategy()
        {
            return OnGetJobPersistenceStrategy();
        }
        public IPersistenceStrategy<JobOrderEntryData> GetJobOrderEntryDataPersistenceStrategy()
        {
            return OnGetJobOrderEntryDataPersistenceStrategy();
        }
        public IPersistenceStrategy<Printer> GetPrinterPersistenceStrategy()
        {
            return OnGetPrinterPersistenceStrategy();
        }
        protected virtual IPersistenceStrategy<Glazing> OnGetGlazingPersistenceStrategy()
        {
            throw new NotImplementedException();
        }
        protected virtual IPersistenceStrategy<Job> OnGetJobPersistenceStrategy()
        {
            throw new NotImplementedException();
        }
        protected virtual IPersistenceStrategy<JobOrderEntryData> OnGetJobOrderEntryDataPersistenceStrategy()
        {
            throw new NotImplementedException();
        }
        protected virtual IPersistenceStrategy<Printer> OnGetPrinterPersistenceStrategy()
        {
            throw new NotImplementedException();
        }
    }
}
