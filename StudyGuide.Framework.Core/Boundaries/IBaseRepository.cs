using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace StudyGuide.Framework.Core.Boundaries
{
    public interface IBaseRepository<T> where T : class
    {
        Task<IEnumerable<T>> AllAsync();

        Task<T> FindAsync(dynamic id);

        IEnumerable<T> Where(Expression<Func<T, bool>> p);

        T Add(T item);

        IEnumerable<T> AddRange(IEnumerable<T> items);

        void Update(T item);

        void UpdateAll(ISet<T> items);

        Task<long> CountAsync();

        Task<long> CountAsync(Expression<Func<T, bool>> p);

        Task<int> SaveAsync();

        void DeleteAll(ISet<T> items);

        void Delete(T item);
    }
}
