using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;
using TriviaExcelGenerator.Models;

namespace TriviaExcelGenerator
{

    public class Program
    {
        private static HttpClient jServiceClient = new HttpClient();
        private static HttpClient openTriviaClient = new HttpClient();
        private static string excelpath = @"D:\TriviaExcel.xlsx";
        private static int MaxQuestions = 50;

        //// http://jservice.io
        public const string jServiceApiUrl = "/api/random";

        //// https://opentdb.com/api_config.php
        public const string openTriviaApiUrl = "/api.php?amount=50&type=multiple";
        
        public static void Main(string[] args)
        {
            // QUIZ API Base Addresses
            jServiceClient.BaseAddress = new Uri("http://jservice.io");
            openTriviaClient.BaseAddress = new Uri("https://opentdb.com");

            // Create a data table which will be exported to excel
            var dataTable = new DataTable();
            dataTable.Columns.Add("Question");
            dataTable.Columns.Add("Correct Answer");
            dataTable.Columns.Add("Incorrect Answer 1");
            dataTable.Columns.Add("Incorrect Answer 2");

            // Create an async task to make API calls
            var task = Task.Run(async () =>
            {
                // Create questions
                for (var i = 0; i < MaxQuestions; i++)
                {
                    Console.WriteLine($"Generating Question {i + 1}");

                    // Create a random number to make call to 2 different APIs
                    var random = new Random();
                    var randomNumber = random.Next();

                    // Get the question and answer from the APIs
                    var result = await GetQuestion(randomNumber);

                    // Create a new row in the data table using the data from API
                    var dataRow = dataTable.NewRow();
                    dataRow[0] = result.Question;
                    dataRow[1] = result.CorrectAnswer;
                    dataRow[2] = result.IncorrectAnswer1;
                    dataRow[3] = result.IncorrectAnswer2;
                    dataTable.Rows.Add(dataRow);
                }
            });

            // Wait till the async task is completed
            Task.WaitAll(task);

            // Generate the excel using the data table
            GenerateExcel(dataTable);
        }

        private static async Task<QuizFormat> GetQuestion(int randomNumber)
        {
            var output = string.Empty;
            var result = new QuizFormat();

            // Make a random call to any of the APIs based on the random number generated
            // The random call is to bypass if the Google Actions performs a check based on this API :)
            if (randomNumber % 2 == 0)
            {
                // Get the response from the API
                output = await jServiceClient.GetStringAsync(jServiceApiUrl);

                // Deserialize the response to the format we need
                var quizFormat = JsonConvert.DeserializeObject<List<JServiceQuizFormat>>(output);
                result.CorrectAnswer = HttpUtility.HtmlDecode(quizFormat[0].answer);
                result.Question = HttpUtility.HtmlDecode(quizFormat[0].question);

                // The API returns only question and answer.
                // Since we need multi-choice make 2 more calls to the API and create two more options :)
                // It wont be always correct. But just works most of the time!!!
                output = await jServiceClient.GetStringAsync(jServiceApiUrl);
                quizFormat = JsonConvert.DeserializeObject<List<JServiceQuizFormat>>(output);
                result.IncorrectAnswer1 = HttpUtility.HtmlDecode(quizFormat[0].answer);

                output = await jServiceClient.GetStringAsync(jServiceApiUrl);
                quizFormat = JsonConvert.DeserializeObject<List<JServiceQuizFormat>>(output);
                result.IncorrectAnswer2 = HttpUtility.HtmlDecode(quizFormat[0].answer);
            }
            else
            {
                // Get the response from the API
                output = await openTriviaClient.GetStringAsync(openTriviaApiUrl);

                // Deserialize the response to the format we need
                var quizFormat = JsonConvert.DeserializeObject<OpenTriviaQuizFormat>(output);
                result.CorrectAnswer = HttpUtility.HtmlDecode(quizFormat.results[0].correct_answer);
                result.Question = HttpUtility.HtmlDecode(quizFormat.results[0].question);
                result.IncorrectAnswer1 = HttpUtility.HtmlDecode(quizFormat.results[0].incorrect_answers[0]);
                result.IncorrectAnswer2 = HttpUtility.HtmlDecode(quizFormat.results[0].incorrect_answers[1]);
            }

            return result;
        }

        public static void GenerateExcel(DataTable dataTable)
        {
            // Create a FileInfo object from the path
            var finame = new FileInfo(excelpath);

            // Delete the file if already exists
            if (File.Exists(excelpath))
            {
                File.Delete(excelpath);
            }

            // If not exists, create it
            if (!File.Exists(excelpath))
            {
                // Create an excel package
                ExcelPackage excel = new ExcelPackage(finame);

                // Add the worksheets
                var sheetcreate = excel.Workbook.Worksheets.Add(DateTime.UtcNow.Date.ToString());

                // Create the rows based on the datatable
                if (dataTable.Rows.Count > 0)
                {
                    // Create the headers
                    sheetcreate.Cells[1, 1].Value = dataTable.Columns[0].ColumnName;
                    sheetcreate.Cells[1, 2].Value = dataTable.Columns[1].ColumnName;
                    sheetcreate.Cells[1, 3].Value = dataTable.Columns[2].ColumnName;
                    sheetcreate.Cells[1, 4].Value = dataTable.Columns[3].ColumnName;

                    // Create the data rows
                    for (int i = 1; i < dataTable.Rows.Count;)
                    {
                        sheetcreate.Cells[i + 1, 1].Value = dataTable.Rows[i][0].ToString();
                        sheetcreate.Cells[i + 1, 2].Value = dataTable.Rows[i][1].ToString();
                        sheetcreate.Cells[i + 1, 3].Value = dataTable.Rows[i][2].ToString();
                        sheetcreate.Cells[i + 1, 4].Value = dataTable.Rows[i][3].ToString();
                        i++;
                    }
                }
                
                // Save the file
                excel.Save();  
            }
        }
    }
}
