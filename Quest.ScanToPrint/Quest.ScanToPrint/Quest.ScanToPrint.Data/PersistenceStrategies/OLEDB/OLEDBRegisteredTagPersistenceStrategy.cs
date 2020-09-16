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
    public class OLEDBRegisteredTagPersistenceStrategy : OLEDBPersistenceStrategy<RegisteredTag>
    {
        public OLEDBRegisteredTagPersistenceStrategy(string dbConnectionString) : base(dbConnectionString) { }
        protected override void DefineAddCommand(ref OleDbCommand command, RegisteredTag entity)
        {
            command.CommandText = "INSERT INTO RegisteredTag([Job],[Floor],[Tag]) VALUES(?,?,?)";

            command.Parameters.AddWithValue("@Job", entity.Job);
            command.Parameters.AddWithValue("@Floor", entity.Floor);
            command.Parameters.AddWithValue("@Tag", entity.Tag);
        }
        protected override void DefineEditCommand(ref OleDbCommand command, RegisteredTag entity)
        {
            command.CommandText = "UPDATE RegisteredTag SET [Job] = ?, [Floor] = ?, [Tag] = ? WHERE [ID] = ?";

            command.Parameters.AddWithValue("@Job", entity.Job);
            command.Parameters.AddWithValue("@Floor", entity.Floor);
            command.Parameters.AddWithValue("@Tag", entity.Tag);
            command.Parameters.AddWithValue("@ID", entity.ID);
        }
        protected override void DefineDeleteCommand(ref OleDbCommand command, RegisteredTag entity)
        {
            command.CommandText = "DELETE FROM RegisteredTag WHERE [ID] = ?";

            command.Parameters.AddWithValue("@ID", entity.ID);
        }
        protected override void DefineSelectCommand(ref OleDbCommand command, QueryParameters predicate)
        {
            command.CommandText = "SELECT [ID],[Job],[Floor],[Tag] FROM RegisteredTag";
            if (predicate != null)
            {
                foreach (string item in predicate.Keys)
                {
                    command.Parameters.AddWithValue(item, predicate[item]);
                }
            }
        }
        protected override RegisteredTag SelectResultItemMapping(DataRow dataRow)
        {
            return new RegisteredTag()
            {
                ID = (int)dataRow["ID"],
                Job = dataRow["Job"].ToString(),
                Floor = dataRow["Floor"].ToString(),
                Tag = dataRow["Tag"].ToString()
            };
        }
    }
}
