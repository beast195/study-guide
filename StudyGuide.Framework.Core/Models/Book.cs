using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudyGuide.Framework.Core.Models
{
    public class Book
    {
        public int Id { get; set; }
        public string BookTitle { get; set; }
        public DateTime LastRead { get; set; }
        public BookType BookType { get; set; }
        public bool BookOpened { get; set; }
        public string BookFileName { get; set; }
        public int ChaptersRead { get; set; }
        public int ChapterTotal { get; set; }
        public virtual ISet<Chapter> Chapters { get; set; }
    }
    public enum BookType
    {
        PDF, WORD
    }
}
