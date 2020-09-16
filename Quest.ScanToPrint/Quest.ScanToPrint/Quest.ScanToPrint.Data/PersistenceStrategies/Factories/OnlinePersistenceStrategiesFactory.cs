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
        public IPersistenceStrategy<Glazing> GetGlazingPersistenceStrategy()
        {
            return OnGetGlazingPersistenceStrategy();
        }
        public IPersistenceStrategy<Job> GetJobPersistenceStrategy()
        {
            return OnGetJobPersistenceStrategy();
        }
        protected virtual IPersistenceStrategy<Glazing> OnGetGlazingPersistenceStrategy()
        {
            throw new NotImplementedException();
        }
        protected virtual IPersistenceStrategy<Job> OnGetJobPersistenceStrategy()
        {
            throw new NotImplementedException();
        }
    }
}
