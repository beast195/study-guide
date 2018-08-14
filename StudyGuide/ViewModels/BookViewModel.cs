using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace StudyGuide.ViewModels
{
    public class BookViewModel
    {
        public int BookId { get; set; }
        public string BookName { get; set; }
        public int NumberOfChapters { get; set; }
        public bool Open { get; set; }

        public List<ChapterViewModel> Chapers { get; set; }
        public string ChapterName { get; set; }
    }

    public class BookUploadViewModel
    {
        public int BookId { get; set; }
        public int NumberOfChapters { get; set; }
    }
}