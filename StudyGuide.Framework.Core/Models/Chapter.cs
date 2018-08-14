using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudyGuide.Framework.Core.Models
{
    public class Chapter
    {
        public int Id { get; set; }
        public Book Book { get; set; }
        public int Number { get; set; }
        public string Content { get; set; }
        public DateTime LastRead { get; set; }
        public bool ChapterOpened { get; set; }
    }
}
