using System.Collections.Generic;

namespace TriviaExcelGenerator.Models
{
    public class OpenTriviaQuestion
    {
        public string question { get; set; }
        public string correct_answer { get; set; }
        public List<string> incorrect_answers { get; set; }
    }
}
