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
    public class OLEDBJobOrderEntryDataPersistenceStrategy : OLEDBPersistenceStrategy<JobOrderEntryData>
    {
        public OLEDBJobOrderEntryDataPersistenceStrategy(string dbConnectionString) : base(dbConnectionString) { }        
        protected override void DefineSelectCommand(ref OleDbCommand command, QueryParameters predicate)
        {
            command.CommandText = "SELECT [Job],[Floor],[Tag] FROM JobOrderEntryData";
            if (predicate != null)
            {
                foreach (string item in predicate.Keys)
                {
                    command.Parameters.AddWithValue(item, predicate[item]);
                }
            }
        }
        protected override JobOrderEntryData SelectResultItemMapping(DataRow dataRow)
        {
            return new JobOrderEntryData()
            {
                Job = dataRow["Job"].ToString(),
                Floor = dataRow["Floor"].ToString(),
                Tag = dataRow["Tag"].ToString()
            };
        }
    }
}
