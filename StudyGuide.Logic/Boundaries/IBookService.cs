using StudyGuide.Framework.Core.Models;
using StudyGuide.Logic.Entities;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace StudyGuide.Logic.Boundaries
{
    public interface IBookService
    {
        Task CreateBook(string filename);
        ChapterEntity GetOpenBook();
        List<BookEntity> GetAllOpenBooks();
        List<BookEntity> GetAllBooks();
        Task OpenNewChapter(Book book, int chapterNumber);
        Task OpenChapterAvailble(int bookId, int chapNumber);
    }
}