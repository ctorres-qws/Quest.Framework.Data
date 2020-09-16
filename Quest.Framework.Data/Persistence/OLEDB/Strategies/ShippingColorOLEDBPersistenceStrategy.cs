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
    public class ShippingColorOLEDBPersistenceStrategy : OLEDBPersistenceStrategy<ShippingColor>
    {
        public ShippingColorOLEDBPersistenceStrategy(string dbConnectionString) : base(dbConnectionString) { }
        protected override void DefineSelectCommand(ref OleDbCommand command, QueryParameters predicate)
        {            
            command.CommandText = "SELECT ID,ColorName,ColorHexRGB FROM X_Shipping_Colors";
        }
        protected override ShippingColor SelectResultItemMapping(DataRow dataRow)
        {
            return new ShippingColor()
            {
                ID = Convert.ToInt32(dataRow["ID"]),
                ColorName = dataRow["ColorName"].ToString(),
                ColorHexRGB = dataRow["ColorHexRGB"].ToString()
            };
        }
    }
}
