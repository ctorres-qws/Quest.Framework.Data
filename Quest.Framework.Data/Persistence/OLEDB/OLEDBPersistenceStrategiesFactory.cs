using Quest.Framework.Data.Entities;
using Quest.Framework.Data.Persistence.OLEDB.Strategies;
using Quest.Framework.Persistance;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quest.Framework.Data.Persistence.OLEDB
{
    public class OLEDBPersistenceStrategiesFactory : PersistenceStrategiesFactory
    {
        public string DbConnectionString { get; set; }
        public OLEDBPersistenceStrategiesFactory(string dbConnectionString)
        {
            this.JobPersistenceStrategy = new JobOLEDBPersistenceStrategy(dbConnectionString);
            this.ShippingColorPersistenceStrategy = new ShippingColorOLEDBPersistenceStrategy(dbConnectionString);
            this.JobShippingColorPersistenceStrategy = new JobShippingColorOLEDBPersistenceStrategy(dbConnectionString);
            this.PrinterPersistenceStrategy = new PrinterOLEDBPersistenceStrategy(dbConnectionString);
            this.UserPersistenceStrategy = new UserOLEDBPersistenceStrategy(dbConnectionString);
        }
        protected override IPersistenceStrategy<Job> OnGetJobPersistenceStrategy()
        {
            return this.JobPersistenceStrategy;
        }
        protected override IPersistenceStrategy<ShippingColor> OnGetShippingColorPersistenceStrategy()
        {
            return this.ShippingColorPersistenceStrategy;
        }
        protected override IPersistenceStrategy<JobShippingColor> OnGetJobShippingColorPersistenceStrategy()
        {
            return this.JobShippingColorPersistenceStrategy;
        }
        protected override IPersistenceStrategy<Printer> OnGetPrinterPersistenceStrategy()
        {
            return this.PrinterPersistenceStrategy;
        }
        protected override IPersistenceStrategy<User> OnGetUserPersistenceStrategy()
        {
            return this.UserPersistenceStrategy;
        }
    }
}
