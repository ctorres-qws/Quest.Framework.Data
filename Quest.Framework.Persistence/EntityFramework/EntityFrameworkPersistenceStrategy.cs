using Quest.Framework.Persistance;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace Quest.Framework.Persistance.EntityFramework
{
    public class PersistenceStrategyEntityFramework<TEntity,TContext> : IPersistenceStrategy<TEntity> 
        where TEntity : class
        where TContext : DbContext
    {

        public void Add(TEntity entity)
        {
            using (DbContext context = (DbContext)Activator.CreateInstance<TContext>())
            {
                context.Set<TEntity>().Add(entity);
                context.SaveChanges();
            }
        }
        public void Edit(TEntity entity) {
            using (DbContext context = (DbContext)Activator.CreateInstance<TContext>())
            {
                context.Entry(entity).State = EntityState.Modified;
                context.SaveChanges();
            }
        }
        public void Delete(TEntity entity)
        {
            using (DbContext context = (DbContext)Activator.CreateInstance<TContext>())
            {
                context.Entry(entity).State = EntityState.Deleted;
                context.SaveChanges();
            }
        }
        public void Delete(Expression<Func<TEntity, bool>> predicate)
        {
            using (DbContext context = (DbContext)Activator.CreateInstance<TContext>())
            {
                TEntity entity = context.Set<TEntity>().Find(predicate);
                context.Entry(entity).State = EntityState.Deleted;
                context.SaveChanges();
            }
        }
        public IEnumerable<TEntity> GetCollection()
        {
            List<TEntity> entities;

            using (DbContext context = (DbContext)Activator.CreateInstance<TContext>())
            {
                entities = context.Set<TEntity>().ToList();
            }

            return entities;
        }
        public IEnumerable<TEntity> GetCollectionFromAJobTable(string tableName)
        {
            throw new NotImplementedException();
        }
        //public IEnumerable<TEntity> GetCollection(Expression<Func<TEntity, bool>> predicate)
        //{
        //    List<TEntity> entities;

        //    using (DbContext context = (DbContext)Activator.CreateInstance<TContext>())
        //    {
        //        entities = context.Set<TEntity>().Where(predicate).ToList();
        //    }

        //    return entities;
        //}
    }
}
