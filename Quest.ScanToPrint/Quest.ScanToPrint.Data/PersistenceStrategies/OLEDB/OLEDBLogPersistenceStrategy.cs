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
    public class OLEDBLogPersistenceStrategy : OLEDBPersistenceStrategy<Log>
    {
        public OLEDBLogPersistenceStrategy(string dbConnectionString) : base(dbConnectionString) { }
        protected override void DefineAddCommand(ref OleDbCommand command, Log entity)
        {
            command.CommandText = "INSERT INTO Log([Description],[DateTime],[ExceptionMessage],[ExceptionStackTrace]) VALUES(?,?,?,?)";

            command.Parameters.AddWithValue("@Description", entity.Description);
            command.Parameters.AddWithValue("@DateTime", DateTime.Parse(entity.DateTime.ToString()));
            command.Parameters.AddWithValue("@ExceptionMessage", entity.ExceptionMessage);
            command.Parameters.AddWithValue("@ExceptionStackTrace", entity.ExceptionStackTrace);
        }
        protected override void DefineSelectCommand(ref OleDbCommand command, QueryParameters predicate)
        {
            command.CommandText = "SELECT [ID],[Description],[DateTime],[ExceptionMessage],[ExceptionStackTrace] FROM Log";
            if(predicate != null)
            {
                foreach(string item in predicate.Keys)
                {
                    command.Parameters.AddWithValue(item, predicate[item]);
                }
            }
        }
        protected override Log SelectResultItemMapping(DataRow dataRow)
        {
            return new Log()
            {
                ID = (int) dataRow["ID"],
                Description = dataRow["Description"].ToString(),
                DateTime = (DateTime) dataRow["DateTime"],
                ExceptionMessage = dataRow["ExceptionMessage"].ToString(),
                ExceptionStackTrace = dataRow["ExceptionStackTrace"].ToString()
            };
        }
    }
}
