using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace SelectQAForMe.Tools
{
    /// <summary>
    ///
    /// </summary>
    public class QAFileParser
    {
        #region Private Fields

        /// <summary>
        /// The opts
        /// </summary>
        private readonly Options _opts;

        /// <summary>
        /// The row items
        /// </summary>
        private List<RowItem> _rowItems;

        /// <summary>
        /// The selected questions
        /// </summary>
        private List<RowItem> _selectedQuestions;

        #endregion Private Fields

        #region Public Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="QAFileParser"/> class.
        /// </summary>
        /// <param name="opts">The opts.</param>
        public QAFileParser(Options opts)
        {
            _opts = opts;
            Console.WriteLine("Error check on input file");
            //TryOpenFile();
            _rowItems = new List<RowItem>();
            _selectedQuestions = new List<RowItem>();
        }

        #endregion Public Constructors

        #region Public Methods

        /// <summary>
        /// Extracts the data to out file.
        /// </summary>
        /// <returns></returns>
        public bool ExtractDataToOutFile()
        {
            try
            {
                using (var stream = new FileStream(_opts.FileName, FileMode.Open))
                {
                    var xssfwb = new XSSFWorkbook(stream);
                    var tabName = string.IsNullOrWhiteSpace(_opts.WorkSheet) ? "Questions" : _opts.WorkSheet;
                    var sheet = xssfwb.GetSheet(tabName);
                    Console.WriteLine("Starting to read file");
                    ExtractDataFromSheet(sheet);
                    Console.WriteLine($"Finished reading input s/s{_opts.FileName}");
                    PickQuestions(_opts.NumberOfQuestions);
                    SaveQuestionsToFile();
                    Console.WriteLine($"Created New file {_opts.OutFileName}");
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"Error while opening file {_opts.FileName}");
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine("Operation failed due to above reason  ):");
                return false;
            }
            return true;
        }

        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Extracts the data from sheet.
        /// </summary>
        /// <param name="sheet">The sheet.</param>
        private void ExtractDataFromSheet(ISheet sheet)
        {
            var rowCount = sheet.LastRowNum;
            rowCount = rowCount >= 230 ? 230 : rowCount;
            for (var i = 0; i < rowCount; i++)
            {
                var row = sheet.GetRow(i);
                var colA = row.GetCell(0);
                var rowItem = new RowItem();
                if (colA.CellType == CellType.Numeric)
                {
                    rowItem.RowNumber = Convert.ToInt32(colA.NumericCellValue);
                }
                else
                {
                    Console.WriteLine("Error reading value");
                    Console.WriteLine($"Sheet {sheet.SheetName}; row: {i + 1}; Column 'A'");
                    Console.WriteLine($"Expected Numeric value; Actual value {colA.CellType}");
                    Environment.Exit(-2);
                }
                var colB = row.GetCell(1);
                if (colB.CellType == CellType.String)
                {
                    string formattedText = RemoveLeadingQuestionNumber(colB.StringCellValue);
                    rowItem.Question = formattedText;
                }
                else
                {
                    Console.WriteLine("Error reading value");
                    Console.WriteLine($"Sheet {sheet.SheetName}; row: {i + 1}; Column 'B'");
                    Console.WriteLine($"Expected String value; Actual value {colA.CellType}");
                    Environment.Exit(-2);
                }
                _rowItems.Add(rowItem);
            }
        }

        /// <summary>
        /// Picks the questions.
        /// </summary>
        /// <param name="optsNumberOfQuestions">The opts number of questions.</param>
        private void PickQuestions(int optsNumberOfQuestions)
        {
            var random = new Random();
            var randomNumbers = new List<int>();
            var rowCount = optsNumberOfQuestions < _rowItems.Count ?
                optsNumberOfQuestions : _rowItems.Count;
            while (randomNumbers.Count < rowCount)
            {
                var rInt = random.Next(0, _rowItems.Count - 2) + 1;
                if (!randomNumbers.Exists(a => a == rInt))
                {
                    randomNumbers.Add(rInt);
                }
            }
            var questCount = randomNumbers.Count;
            for (var i = 0; i < questCount; i++)
            {
                _selectedQuestions.Add(_rowItems[randomNumbers[i]]);
            }
        }

        /// <summary>
        /// Removes the leading question number.
        /// </summary>
        /// <param name="colBStringCellValue">The col b string cell value.</param>
        /// <returns></returns>
        private string RemoveLeadingQuestionNumber(string colBStringCellValue)
        {
            var pattern = @"(^\d+\.\s)";
            var substitution = @"";
            var options = RegexOptions.Multiline;
            var regex = new Regex(pattern, options);
            var result = regex.Replace(colBStringCellValue, substitution);
            return result;
        }

        /// <summary>
        /// Saves the questions to file.
        /// </summary>
        private void SaveQuestionsToFile()
        {
            try
            {
                using (var writer = File.CreateText(_opts.OutFileName))
                {
                    foreach (var selectedQuestion in _selectedQuestions)
                    {
                        writer.WriteLine(selectedQuestion.Question);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error while writing to output file");
                Console.WriteLine($"Reason {e.Message}");
                Environment.Exit(-3);
            }
        }

        /// <summary>
        /// Tries the open file.
        /// </summary>
        private void TryOpenFile()
        {
            try
            {
                using (var stream = new FileStream(_opts.FileName, FileMode.Open))
                {
                    var xssfwb = new XSSFWorkbook(stream);
                    var tabName = string.IsNullOrWhiteSpace(_opts.WorkSheet) ? "Questions" : _opts.WorkSheet;
                    var sheet = xssfwb.GetSheet(tabName);
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"Error while opening file {_opts.FileName}");
                Console.WriteLine($"Error: {ex.Message}");
                Environment.Exit(-1);
            }
        }

        #endregion Private Methods
    }
}