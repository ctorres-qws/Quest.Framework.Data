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
    public class OLEDBJobShippingLabelColorPersistenceStrategy : OLEDBPersistenceStrategy<JobShippingLabelColor>
    {
        public OLEDBJobShippingLabelColorPersistenceStrategy(string dbConnectionString) : base(dbConnectionString) { }
        protected override void DefineAddCommand(ref OleDbCommand command, JobShippingLabelColor entity)
        {
            command.CommandText = "INSERT INTO JobShippingLabelColor([Job],[Color]) VALUES(?,?)";

            command.Parameters.AddWithValue("@Job", entity.Job);
            command.Parameters.AddWithValue("@Color", entity.Color);
        }
        protected override void DefineEditCommand(ref OleDbCommand command, JobShippingLabelColor entity)
        {
            command.CommandText = "UPDATE JobShippingLabelColor SET [Color] = ? WHERE [Job] = ?";

            command.Parameters.AddWithValue("@Color", entity.Color);
            command.Parameters.AddWithValue("@Job", entity.Job);
        }
        protected override void DefineDeleteCommand(ref OleDbCommand command, JobShippingLabelColor entity)
        {
            command.CommandText = "DELETE FROM JobShippingLabelColor WHERE [Job] = ?";

            command.Parameters.AddWithValue("@Job", entity.Job);
        }
        protected override void DefineSelectCommand(ref OleDbCommand command, QueryParameters predicate)
        {
            command.CommandText = "SELECT [Job],[Color] FROM JobShippingLabelColor";
            if(predicate != null)
            {
                foreach(string item in predicate.Keys)
                {
                    command.Parameters.AddWithValue(item, predicate[item]);
                }
            }
        }
        protected override JobShippingLabelColor SelectResultItemMapping(DataRow dataRow)
        {
            return new JobShippingLabelColor()
            {
                Color = dataRow["Color"].ToString(),
                Job = dataRow["Job"].ToString()
            };
        }
    }
}
