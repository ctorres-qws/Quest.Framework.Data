using Quest.Framework.Persistance;
using Quest.ScanToPrint.Data.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quest.ScanToPrint.Data
{
    public abstract class LocalPersistenceStrategiesFactory : IPersistenceStrategiesFactory
    {
        internal IPersistenceStrategy<Barcodes> BarcodesPersistenceStrategy { get; set; }
        internal IPersistenceStrategy<JobShippingLabelColor> JobShippingLabelColorPersistenceStrategy { get; set; }
        internal IPersistenceStrategy<Log> LogPersistenceStrategy { get; set; }
        internal IPersistenceStrategy<RegisteredTag> RegisteredTagPersistenceStrategy { get; set; }
        internal IPersistenceStrategy<Printer> PrinterPersistenceStrategy { get; set; }

        public IPersistenceStrategy<Barcodes> GetBarcodesPersistanceStrategy()
        {
            return OnGetBarcodesPersistanceStrategy();
        }
        public IPersistenceStrategy<JobShippingLabelColor> GetJobShippingLabelColorPersistenceStrategy()
        {
            return OnGetJobShippingLabelColorPersistenceStrategy();
        }
        public IPersistenceStrategy<Log> GetLogPersistenceStrategy()
        {
            return OnGetLogPersistenceStrategy();
        }
        public IPersistenceStrategy<RegisteredTag> GetRegisteredTagPersistenceStrategy()
        {
            return OnGetRegisteredTagPersistenceStrategy();
        }
        public IPersistenceStrategy<Printer> GetPrinterPersistenceStrategy()
        {
            return OnGetPrinterPersistenceStrategy();
        }
        protected virtual IPersistenceStrategy<Barcodes> OnGetBarcodesPersistanceStrategy()
        {
            throw new NotImplementedException();
        }
        protected virtual IPersistenceStrategy<JobShippingLabelColor> OnGetJobShippingLabelColorPersistenceStrategy()
        {
            throw new NotImplementedException();
        }
        protected virtual IPersistenceStrategy<Log> OnGetLogPersistenceStrategy()
        {
            throw new NotImplementedException();
        }
        protected virtual IPersistenceStrategy<RegisteredTag> OnGetRegisteredTagPersistenceStrategy()
        {
            throw new NotImplementedException();
        }
        protected virtual IPersistenceStrategy<Printer> OnGetPrinterPersistenceStrategy()
        {
            throw new NotImplementedException();
        }
    }
}
