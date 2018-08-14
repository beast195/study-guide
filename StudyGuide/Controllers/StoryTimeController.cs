using Google.Apis.Services;
using Google.Apis.YouTube.v3;
using StudyGuide.Logic.Boundaries;
using StudyGuide.Logic.Services;
using StudyGuide.ViewModels;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace StudyGuide.Controllers
{
    public class StoryTimeController : Controller
    {
        private readonly IBookService _bookService;

        public StoryTimeController(BookService bookService)
        {
            _bookService = bookService;
        }

        // GET: StoryTime
        public async Task<ActionResult> Index()
        {
            var youtubeService = new YouTubeService(new BaseClientService.Initializer()
            {
                ApiKey = ConfigurationManager.AppSettings["google.api.key"].ToString(),
                ApplicationName = this.GetType().ToString()
            });

            var videos = new List<StoryTimeViewModel>();
            var searchTerms = _bookService.GetAllOpenBooks().SelectMany(c => c.Chapers).Select(c => c.Content).FirstOrDefault().Split(new char[] { ',', '.', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            // Add each result to the appropriate list, and then display the lists of
            // matching videos, channels, and playlists.

            var searchListRequest = youtubeService.Search.List("snippet");
            foreach (var searchTerm in searchTerms)
            {
                if (!String.IsNullOrWhiteSpace(searchTerm) && searchTerm.Length > 3)
                {
                    searchListRequest.Q = searchTerm; // Replace with your search term.
                    searchListRequest.MaxResults = 3;
                    
                        // Call the search.list method to retrieve results matching the specified query term.
                        var searchListResponse = await searchListRequest.ExecuteAsync();
                        foreach (var searchResult in searchListResponse.Items)
                        {
                            if (searchResult.Id.Kind == "youtube#video")
                            {
                                videos.Add(new StoryTimeViewModel() { LinkTitle = searchResult.Snippet.Title, Link = $"https://youtube.com/embed/"+searchResult.Id.VideoId });
                            }

                            break;
                        }
                    
                }
            }

            return View(videos);
        }

        // GET: StoryTime/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: StoryTime/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: StoryTime/Create
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

        // GET: StoryTime/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        // POST: StoryTime/Edit/5
        [HttpPost]
        public ActionResult Edit(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add update logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        // GET: StoryTime/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: StoryTime/Delete/5
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