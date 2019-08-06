using System.Collections.Generic;

namespace TriviaExcelGenerator.Models
{
    public class OpenTriviaQuizFormat
    {
        public int response_code { get; set; }

        public List<OpenTriviaQuestion> results { get; set; } 
    }
}
