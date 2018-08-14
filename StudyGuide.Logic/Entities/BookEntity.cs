using System.Collections.Generic;

namespace StudyGuide.Logic.Entities
{
    public class BookEntity
    {
        public int BookId { get; set; }
        public string BookName { get; set; }
        public int NumberOfChapters { get; set; }
        public bool Open { get; set; }

        public List<ChapterEntity> Chapers { get; set; }
    }
}