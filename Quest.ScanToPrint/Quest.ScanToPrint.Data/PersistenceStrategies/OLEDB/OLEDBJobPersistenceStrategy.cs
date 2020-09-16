using Quest.Framework.Persistance;
using Quest.Framework.Persistance.OLEDB;
using Quest.ScanToPrint.Data.Entities;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quest.ScanToPrint.Data
{
    public class OLEDBJobPersistenceStrategy : OLEDBPersistenceStrategy<Job>
    {
        public OLEDBJobPersistenceStrategy(string dbConnectionString) : base(dbConnectionString) { }
        protected override void DefineSelectCommand(ref OleDbCommand command, QueryParameters predicate)
        {
            command.CommandText = "SELECT [JOB],[ShippingLabelColor] FROM Z_Jobs";
            if (predicate != null)
            {
                foreach (string item in predicate.Keys)
                {
                    command.Parameters.AddWithValue(item, predicate[item]);
                }
            }
        }
        protected override Job SelectResultItemMapping(DataRow dataRow)
        {
            return new Job()
            {
                JOB = dataRow["JOB"].ToString(),
                ShippingLabelColor = dataRow["ShippingLabelColor"].ToString()
            };
        }
    }
}
