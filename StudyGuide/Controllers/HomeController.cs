using StudyGuide.Logic.Boundaries;
using StudyGuide.Logic.Services;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace StudyGuide.Controllers
{
    public class HomeController : Controller
    {
        private readonly IBookService _bookService;

        public HomeController(BookService bookService)
        {
            _bookService = bookService;
        }
        public HomeController()
        {

        }

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        [HttpPost]
        public async Task<ActionResult> CreateBook(HttpPostedFileWrapper book)
        {
            if (book != null)
            {
                book.SaveAs(Server.MapPath($"/BookFiles/{book.FileName}"));
                await _bookService.CreateBook(Server.MapPath($"/BookFiles/{book.FileName}"));
            }
            return Redirect("/Home/Index");
        }
    }
}