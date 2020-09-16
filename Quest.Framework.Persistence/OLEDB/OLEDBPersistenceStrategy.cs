using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quest.Framework.Persistance.OLEDB
{
    public abstract class OLEDBPersistenceStrategy<TEntity> : IPersistenceStrategy<TEntity> where TEntity : class
    {
        string DbConnectionString { get; set; }
        public OLEDBPersistenceStrategy(string dbConnectionString)
        {
            this.DbConnectionString = dbConnectionString;
        }
        public virtual void Add(TEntity entity)
        {
            using (OleDbConnection connection = new OleDbConnection(DbConnectionString))
            {
                using (OleDbCommand command = CreateAddCommand(connection, entity))
                {
                    connection.Open();

                    command.ExecuteNonQuery();

                    connection.Close();
                }
            }
        }
        public virtual void Edit(TEntity entity)
        {
            using (OleDbConnection connection = new OleDbConnection(DbConnectionString))
            {
                using (OleDbCommand command = CreateEditCommand(connection, entity))
                {
                    connection.Open();

                    command.ExecuteNonQuery();

                    connection.Close();
                }
            }
        }
        public virtual void Delete(TEntity entity)
        {
            using (OleDbConnection connection = new OleDbConnection(DbConnectionString))
            {
                using (OleDbCommand command = CreateDeleteCommand(connection, entity))
                {
                    connection.Open();

                    command.ExecuteNonQuery();

                    connection.Close();
                }
            }
        }
        public virtual IEnumerable<TEntity> GetCollection() { return GetCollectionEntities(); }
        public virtual IEnumerable<TEntity> GetCollectionFromAJobTable(string tableName) { return GetCollectionEntitiesFromAJobTable(tableName); }
        public virtual IEnumerable<TEntity> GetCollectionEntitiesFromAJobTable(string tableName, QueryParameters predicate = null)
        {
            using (OleDbConnection connection = new OleDbConnection(DbConnectionString))
            {
                using (OleDbCommand command = CreateSelectFromAJobTableCommand(connection, tableName, predicate))
                {
                    connection.Open();

                    DataTable table = new DataTable();
                    OleDbDataAdapter adapter = new OleDbDataAdapter();
                    adapter.SelectCommand = command;
                    adapter.Fill(table);

                    foreach (DataRow row in table.Rows)
                    {
                        yield return SelectResultItemMapping(row);
                    }

                    connection.Close();
                }
            }
        }
        public virtual IEnumerable<TEntity> GetCollectionEntities(QueryParameters predicate = null)
        {            
            using (OleDbConnection connection = new OleDbConnection(DbConnectionString))
            {
                using (OleDbCommand command = CreateSelectCommand(connection, predicate))
                {
                    connection.Open();

                    DataTable table = new DataTable();
                    OleDbDataAdapter adapter = new OleDbDataAdapter();
                    adapter.SelectCommand = command;
                    adapter.Fill(table);

                    foreach (DataRow row in table.Rows)
                    {
                        yield return SelectResultItemMapping(row);
                    }

                    connection.Close();
                }
            }
        }        
        public OleDbCommand CreateAddCommand(OleDbConnection connection, TEntity entity)
        {
            OleDbCommand command = new OleDbCommand();
            command.Connection = connection;
            DefineAddCommand(ref command, entity);

            return command;
        }
        
        public OleDbCommand CreateEditCommand(OleDbConnection connection, TEntity entity)
        {
            OleDbCommand command = new OleDbCommand();
            command.Connection = connection;
            DefineEditCommand(ref command, entity);

            return command;
        }
        
        public OleDbCommand CreateDeleteCommand(OleDbConnection connection, TEntity entity)
        {
            OleDbCommand command = new OleDbCommand();
            command.Connection = connection;
            DefineDeleteCommand(ref command, entity);

            return command;
        }
        public OleDbCommand CreateSelectCommand(OleDbConnection connection, QueryParameters predicate)
        {
            OleDbCommand command = new OleDbCommand();
            command.Connection = connection;
            DefineSelectCommand(ref command, predicate);

            return command;
        }
        public OleDbCommand CreateSelectFromAJobTableCommand(OleDbConnection connection, string tableName, QueryParameters predicate)
        {
            OleDbCommand command = new OleDbCommand();
            command.Connection = connection;
            DefineSelectFromAJobTableCommand(ref command, tableName, predicate);

            return command;
        }
        protected virtual void DefineAddCommand(ref OleDbCommand command, TEntity entity)
        {
            throw new NotImplementedException();
        }
        protected virtual void DefineEditCommand(ref OleDbCommand command, TEntity entity)
        {
            throw new NotImplementedException();
        }
        protected virtual void DefineDeleteCommand(ref OleDbCommand command, TEntity entity)
        {
            throw new NotImplementedException();
        }
        protected virtual void DefineSelectCommand(ref OleDbCommand command, QueryParameters predicate)
        {
            throw new NotImplementedException();
        }
        protected virtual void DefineSelectFromAJobTableCommand(ref OleDbCommand command, string tableName, QueryParameters predicate)
        {
            throw new NotImplementedException();
        }
        protected virtual TEntity SelectResultItemMapping(DataRow dataRow)
        {
            throw new NotImplementedException();
        }
    }
}
