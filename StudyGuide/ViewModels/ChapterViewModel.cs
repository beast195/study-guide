using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace StudyGuide.ViewModels
{
    public class ChapterViewModel
    {
        public int Id { get; set; }
        public string Content { get; set; }
        public string BookName { get; set; }
        public int ChapterNumber { get; set; }
        public bool OpenChap { get; set; }
    }
}