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

namespace Quest.Framework.Data.Persistence.OLEDB.Strategies
{
    public class PrinterOLEDBPersistenceStrategy : OLEDBPersistenceStrategy<Printer>
    {
        public PrinterOLEDBPersistenceStrategy(string dbConnectionString) : base(dbConnectionString) { }
        protected override void DefineSelectCommand(ref OleDbCommand command, QueryParameters predicate)
        {
            command.CommandText = "SELECT ID, GlazingLine, Active FROM GlazingLinesPrinters";
        }
        protected override void DefineEditCommand(ref OleDbCommand command, Printer entity)
        {
            command.CommandText = string.Format("UPDATE GlazingLinesPrinters SET Active = {0} WHERE ID = {1}", entity.Active, entity.ID);
        }
        protected override Printer SelectResultItemMapping(DataRow dataRow)
        {
            return new Printer()
            {
                Active = Convert.ToBoolean(dataRow["Active"]),
                GlazingLine = Convert.ToInt32(dataRow["GlazingLine"]),
                ID = Convert.ToInt32(dataRow["ID"])
            };
        }
    }
}
