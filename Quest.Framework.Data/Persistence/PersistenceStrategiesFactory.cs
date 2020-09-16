using Quest.Framework.Data.Entities;
using Quest.Framework.Persistance;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quest.Framework.Data.Persistence
{
    public abstract class PersistenceStrategiesFactory : IPersistenceStrategiesFactory
    {
        internal IPersistenceStrategy<Job> JobPersistenceStrategy { get; set; }
        internal IPersistenceStrategy<ShippingColor> ShippingColorPersistenceStrategy { get; set; }
        internal IPersistenceStrategy<JobShippingColor> JobShippingColorPersistenceStrategy { get; set; }
        internal IPersistenceStrategy<Printer> PrinterPersistenceStrategy { get; set; }
        internal IPersistenceStrategy<User> UserPersistenceStrategy { get; set; }
        public IPersistenceStrategy<Job> GetJobPersistenceStrategy()
        {
            return OnGetJobPersistenceStrategy();
        }
        public IPersistenceStrategy<ShippingColor> GetShippingColorPersistenceStrategy()
        {
            return OnGetShippingColorPersistenceStrategy();
        }
        public IPersistenceStrategy<JobShippingColor> GetJobShippingColorPersistenceStrategy()
        {
            return OnGetJobShippingColorPersistenceStrategy();
        }
        public IPersistenceStrategy<Printer> GetPrinterPersistenceStrategy()
        {
            return OnGetPrinterPersistenceStrategy();
        }
        public IPersistenceStrategy<User> GetUserPersistenceStrategy()
        {
            return OnGetUserPersistenceStrategy();
        }
        protected virtual IPersistenceStrategy<Job> OnGetJobPersistenceStrategy()
        {
            throw new NotImplementedException();
        }
        protected virtual IPersistenceStrategy<ShippingColor> OnGetShippingColorPersistenceStrategy()
        {
            throw new NotImplementedException();
        }
        protected virtual IPersistenceStrategy<JobShippingColor> OnGetJobShippingColorPersistenceStrategy()
        {
            throw new NotImplementedException();
        }
        protected virtual IPersistenceStrategy<Printer> OnGetPrinterPersistenceStrategy()
        {
            throw new NotImplementedException();
        }
        protected virtual IPersistenceStrategy<User> OnGetUserPersistenceStrategy()
        {
            throw new NotImplementedException();
        }
    }
}
