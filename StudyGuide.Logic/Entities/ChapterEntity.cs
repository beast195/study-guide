using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudyGuide.Logic.Entities
{
    public class ChapterEntity
    {
        public int Id { get; set; }
        public string Content { get; set; }
        public string BookName { get; set; }
        public int ChapterNumber { get; set; }
        public bool OpenChap { get; set; }
    }
}
