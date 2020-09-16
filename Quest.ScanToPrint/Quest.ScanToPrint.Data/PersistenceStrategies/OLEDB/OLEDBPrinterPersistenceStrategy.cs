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
    public class OLEDBPrinterPersistenceStrategy : OLEDBPersistenceStrategy<Printer>
    {
        public OLEDBPrinterPersistenceStrategy(string dbConnectionString) : base(dbConnectionString) { }
        protected override void DefineSelectCommand(ref OleDbCommand command, QueryParameters predicate)
        {
            command.CommandText = "SELECT ID, GlazingLine, Active, BackupPrinter FROM GlazingLinesPrinters";
        }
        protected override void DefineEditCommand(ref OleDbCommand command, Printer entity)
        {
            command.CommandText = string.Format("UPDATE GlazingLinesPrinters SET Active = {0}, BackupPrinter = {1} WHERE GlazingLine = {2}", entity.Active, entity.BackupPrinter, entity.GlazingLine);
        }
        protected override Printer SelectResultItemMapping(DataRow dataRow)
        {
            return new Printer()
            {
                Active = Convert.ToBoolean(dataRow["Active"]),
                BackupPrinter = Convert.ToInt32(dataRow["BackupPrinter"]),
                GlazingLine = Convert.ToInt32(dataRow["GlazingLine"]),
                ID = Convert.ToInt32(dataRow["ID"])                
            };
        }
    }
}
