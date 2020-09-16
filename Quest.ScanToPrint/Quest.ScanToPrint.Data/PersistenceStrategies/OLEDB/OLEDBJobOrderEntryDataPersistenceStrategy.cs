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
        protected override void DefineSelectFromAJobTableCommand(ref OleDbCommand command, string tableName, QueryParameters predicate)
        {
            command.CommandText = string.Format("SELECT DISTINCT [Job],[Floor],[Tag],Style,IIF(O1 = \"SW\" Or O2 = \"SW\" Or O3 = \"SW\" Or O4 = \"SW\" Or O5 = \"SW\" Or O6 = \"SW\" Or O7 = \"SW\" Or O8 = \"SW\", 1, 0) AS ContainsSW FROM {0} job INNER JOIN Styles styles on job.Style = CStr(styles.Name)", tableName);
        }
        protected override JobOrderEntryData SelectResultItemMapping(DataRow dataRow)
        {
            return new JobOrderEntryData()
            {
                Job = dataRow["Job"].ToString(),
                Floor = dataRow["Floor"].ToString(),
                Tag = dataRow["Tag"].ToString(),
                ContainsSW = Convert.ToBoolean(dataRow["ContainsSW"])
            };
        }
    }
}
