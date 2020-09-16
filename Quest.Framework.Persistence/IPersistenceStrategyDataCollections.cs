using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace Quest.Framework.Persistance
{
    public interface IPersistenceStrategyDataCollections<TEntity>
    {
        void AddCollection(TEntity entity);
        void EditCollection(TEntity entity);
        void DeleteCollection(TEntity entity);
    }
}
