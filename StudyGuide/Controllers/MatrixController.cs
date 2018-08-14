using StudyGuide.Logic.Boundaries;
using StudyGuide.Logic.Services;
using StudyGuide.ViewModels;
using System.Web.Mvc;

namespace StudyGuide.Controllers
{
    public class MatrixController : Controller
    {
        private readonly IBookService _bookService;

        public MatrixController(BookService bookService)
        {
            _bookService = bookService;
        }

        // GET: Matrix
        public ActionResult Index()
        {
            var chapter = new ChapterViewModel();
            var chapterEntity = _bookService.GetOpenBook();
            chapter.Content = chapterEntity.Content;
            chapter.BookName = chapterEntity.BookName;
            return View(chapter);
        }
    }
}