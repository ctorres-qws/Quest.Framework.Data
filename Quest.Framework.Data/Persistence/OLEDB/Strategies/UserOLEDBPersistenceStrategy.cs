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
    public class UserOLEDBPersistenceStrategy : OLEDBPersistenceStrategy<User>
    {
        public UserOLEDBPersistenceStrategy(string dbConnectionString) : base(dbConnectionString) { }
        protected override void DefineSelectCommand(ref OleDbCommand command, QueryParameters predicate)
        {
            command.CommandText = "SELECT Username, Password, FirstName, LastName FROM Users";
        }
        protected override User SelectResultItemMapping(DataRow dataRow)
        {
            return new User()
            {
                Username = dataRow["Username"].ToString(),
                Password = dataRow["Password"].ToString(),
                FirstName = dataRow["FirstName"].ToString(),
                LastName = dataRow["LastName"].ToString()
            };
        }
    }
}
