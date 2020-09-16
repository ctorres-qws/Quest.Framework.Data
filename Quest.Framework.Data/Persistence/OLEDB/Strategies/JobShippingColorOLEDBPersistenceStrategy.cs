using Quest.Framework.Data.Entities;
using Quest.Framework.Persistance;
using Quest.Framework.Persistance.OLEDB;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quest.Framework.Data.Persistence.OLEDB
{
    public class JobShippingColorOLEDBPersistenceStrategy : OLEDBPersistenceStrategy<JobShippingColor>
    {
        public JobShippingColorOLEDBPersistenceStrategy(string dbConnectionString) : base(dbConnectionString) { }
        protected override void DefineSelectCommand(ref OleDbCommand command, QueryParameters predicate)
        {            
            command.CommandText = "SELECT JOB,ShippingLabelColor,ColorName FROM JobShippingColors";
        }
        protected override JobShippingColor SelectResultItemMapping(DataRow dataRow)
        {
            return new JobShippingColor()
            {
                Job = dataRow["JOB"].ToString(),
                ShippingLabelColor = dataRow["ShippingLabelColor"].ToString(),
                ColorName = dataRow["ColorName"].ToString()
            };
        }
    }
}
