using StudyGuide.Framework.Core.Boundaries;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Linq.Expressions;
using System.Data.Entity.Infrastructure;
using System.Data.Entity;

namespace StudyGuide.Framework.Repositories
{
    public class BaseRepository<T> : IBaseRepository<T> where T : class
    {
        protected readonly ApplicationDbContext Context;

        public BaseRepository(ApplicationDbContext context)
        {
            Context = context;
        }

        public T Add(T item)
        {
            AsSet().Add(item);
            return item;
        }

        public IEnumerable<T> AddRange(IEnumerable<T> items)
        {
            AsSet().AddRange(items);
            return items;
        }

        public virtual async Task<IEnumerable<T>> AllAsync()
        {
            return await AsSet().ToListAsync();
        }

        public Task<long> CountAsync()
        {
            return AsSet().LongCountAsync();
        }

        public Task<long> CountAsync(Expression<Func<T, bool>> p)
        {
            return AsSet().LongCountAsync(p);
        }

        public void Delete(T item)
        {
            AsSet().Remove(item);
        }

        public void DeleteAll(ISet<T> items)
        {
            AsSet().RemoveRange(items);
        }

        public virtual async Task<T> FindAsync(dynamic id)
        {
            return await AsSet().FindAsync(id);
        }

        public virtual IEnumerable<T> Paginate<TKey>(int pageSize, int pageNumber, Func<T, TKey> keySelector)
        {
            return AsSet().OrderBy(keySelector)
                          .Skip(pageNumber * pageSize)
                          .Take(pageSize);
        }

        public async Task<int> SaveAsync()
        {
            return await Context.SaveChangesAsync();
        }

        public void Update(T item)
        {
            var entry = item as DbEntityEntry<T>;
            if (entry != null)
            {
                entry.State = EntityState.Modified;
            }
        }

        public void UpdateAll(ISet<T> items)
        {
            var entry = items as DbCollectionEntry<T,T>;
            if (entry != null)
            {
                entry.EntityEntry.State = EntityState.Modified;
            }
        }

        public virtual IEnumerable<T> Where(Expression<Func<T, bool>> p)
        {
            return AsSet().Where(p);
        }

        protected DbSet<T> AsSet()
        {
            return Context.Set<T>();
        }
    }
}