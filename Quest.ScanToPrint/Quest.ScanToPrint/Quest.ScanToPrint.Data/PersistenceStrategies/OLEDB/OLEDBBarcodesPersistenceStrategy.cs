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
    public class OLEDBBarcodesPersistenceStrategy : OLEDBPersistenceStrategy<Barcodes>
    {
        public OLEDBBarcodesPersistenceStrategy(string dbConnectionString) : base(dbConnectionString) { }
        protected override void DefineAddCommand(ref OleDbCommand command, Barcodes entity)
        {
            command.CommandText = "INSERT INTO Barcodes([Job],[Barcode],[ScanDate],[Line],[SentPrint],[SentDatabase],[Tag]) VALUES(?,?,?,?,?,?,?)";

            command.Parameters.AddWithValue("@Job", entity.Job);
            command.Parameters.AddWithValue("@Barcode", entity.Barcode);
            command.Parameters.AddWithValue("@ScanDate", DateTime.Parse(entity.ScanDate.ToString()));
            command.Parameters.AddWithValue("@Line", entity.Line);
            command.Parameters.AddWithValue("@SentPrint", entity.SentPrint);
            command.Parameters.AddWithValue("@SentDatabase", entity.SentDatabase);
            command.Parameters.AddWithValue("@Tag", entity.Tag);
        }
        protected override void DefineEditCommand(ref OleDbCommand command, Barcodes entity)
        {
            command.CommandText = "UPDATE Barcodes SET [Job] = ?, [Barcode] = ?, [ScanDate] = ?, [Line] = ?, [SentPrint] = ?, [SentDatabase] = ?, [Tag] = ? WHERE [ID] = ?";

            command.Parameters.AddWithValue("@Job", entity.Job);
            command.Parameters.AddWithValue("@Barcode", entity.Barcode);
            command.Parameters.AddWithValue("@ScanDate", entity.ScanDate);
            command.Parameters.AddWithValue("@Line", entity.Line);
            command.Parameters.AddWithValue("@SentPrint", entity.SentPrint);
            command.Parameters.AddWithValue("@SentDatabase", entity.SentDatabase);            
            command.Parameters.AddWithValue("@Tag", entity.Tag);
            command.Parameters.AddWithValue("@ID", entity.ID);
        }
        protected override void DefineDeleteCommand(ref OleDbCommand command, Barcodes entity)
        {
            command.CommandText = "DELETE FROM Barcodes WHERE [ID] = ?";

            command.Parameters.AddWithValue("@ID", entity.Job);
        }
        protected override void DefineSelectCommand(ref OleDbCommand command, QueryParameters predicate)
        {
            command.CommandText = "SELECT [ID],[Job],[Barcode],[ScanDate],[Line],[SentPrint],[SentDatabase],[Tag] FROM Barcodes";
            if (predicate != null)
            {
                foreach (string item in predicate.Keys)
                {
                    command.Parameters.AddWithValue(item, predicate[item]);
                }
            }
        }
        protected override Barcodes SelectResultItemMapping(DataRow dataRow)
        {
            return new Barcodes()
            {
                Barcode = dataRow["Barcode"].ToString(),
                ID = (int)dataRow["ID"],
                Job = dataRow["Job"].ToString(),
                Line = (int)dataRow["Line"],
                ScanDate = (DateTime)dataRow["ScanDate"],
                SentDatabase = (bool)dataRow["SentDatabase"],
                SentPrint = (bool)dataRow["SentPrint"],
                Tag = dataRow["Tag"].ToString()
            };
        }
    }
}
