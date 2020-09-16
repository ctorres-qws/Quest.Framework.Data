using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace Quest.Framework.Persistance
{
    public interface IPersistenceStrategy<TEntity>
    {
        void Add(TEntity entity);
        void Edit(TEntity entity);
        void Delete(TEntity entity);
        IEnumerable<TEntity> GetCollection();
        IEnumerable<TEntity> GetCollectionFromAJobTable(string tableName);
    }
}
