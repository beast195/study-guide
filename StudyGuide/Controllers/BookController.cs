using StudyGuide.Logic.Boundaries;
using StudyGuide.Logic.Services;
using StudyGuide.ViewModels;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace StudyGuide.Controllers
{
    public class BookController : Controller
    {
        private readonly IBookService _bookService;

        public BookController(BookService bookService)
        {
            _bookService = bookService;
        }

        // GET: Book
        public ActionResult Index()
        {
            return View(_bookService.GetAllBooks().Select(c => new BookViewModel
            {
                BookId = c.BookId,
                BookName = c.BookName,
                NumberOfChapters = c.NumberOfChapters,
                Open = c.Open,
                Chapers = c.Chapers.Select(s => new ChapterViewModel
                {
                    BookName = s.BookName,
                    Content = s.Content,
                    Id = s.Id,
                    ChapterNumber = s.ChapterNumber,
                    OpenChap = s.OpenChap
                }).ToList(),
                ChapterName = $"Chapter {c.Chapers.Where(j=>j.OpenChap).Select(s => s.ChapterNumber).FirstOrDefault()}"
            }));
        }

        // GET: Book/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: Book/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: Book/Create
        [HttpPost]
        public ActionResult Create(FormCollection collection)
        {
            try
            {
                // TODO: Add insert logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        // GET: Book/Edit/5
        public ActionResult Edit()
        {
            return View();
        }

        // POST: Book/Edit/5
        [HttpPost]
        public async Task<ActionResult> Edit([Bind(Include = "BookId,NumberOfChapters")] int BookId , int NumberOfChapters)
        {
            await _bookService.OpenChapterAvailble(BookId, NumberOfChapters);
            return RedirectToAction("Index");
        }

        // GET: Book/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: Book/Delete/5
        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add delete logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
    }
}