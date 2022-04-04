using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.XlsIO;

namespace LazyDoc
{
    public class Utils
    {
        // print a message to the console
        public void Print(string? message = "Message :")
        {
            // get current time and date
            var now = DateTime.Now;
            var fullMessage = $"{now:dd/MM/yyyy HH:mm:ss.fff} : {message}";
            Console.Write(fullMessage);
        }

        // print message with new line
        public void PrintLine(string message = "Message :")
        {
            // get current time and date
            var now = DateTime.Now;
            var fullMessage = $"{now:dd/MM/yyyy HH:mm:ss.fff} : {message}";
            Console.WriteLine(fullMessage);
        }

        // input a string and return a string
        public string Input(string? msg)
        {
            Print(msg);
            var readLine = Console.ReadLine();
            return (string.IsNullOrEmpty(readLine) ? null : readLine) ?? string.Empty;
        }

        // check file exist and delete
        public void CheckFileExistAndDelete(string filePath)
        {
            if (File.Exists(filePath))
            {
                PrintLine("File already exist.");
                File.Delete(filePath);
                PrintLine("File deleted.");
            }
        }

        public int CheckFileExist(string filePath)
        {
            if (File.Exists(filePath))
            {
                PrintLine("File is already.");
                return 0;
            }
            else
            {
                PrintLine("File not exist.");
                return 1;
            }
        }

        public void ReplaceTextDoc(WordDocument document, Dictionary<string, string> data)
        {
            foreach (var item in data)
            {
                document.Replace(item.Key, item.Value, true, true);
            }
        }

        // get stream from file path
        public Stream GetStreamFromFile(string filePath)
        {
            var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            return stream;
        }

        // create a new file from document stream
        public void CreateFileFromStream(Dictionary<string,string> data, string templateWordPath,string outputWordPath)
        {
            // Creates new Word document instance for word processing.
            using WordDocument subDocument = new WordDocument();
            //Saves the Word document.
            Stream docStream = File.OpenRead(Path.GetFullPath(templateWordPath));
            // open the template document
            subDocument.Open(docStream, FormatType.Doc);
            //Saves the resultant file in the given path.
            docStream = File.Create(Path.GetFullPath(outputWordPath));
            // Finds all occurrences of a misspelled word and replaces with properly spelled word.
            ReplaceTextDoc(subDocument, data);
            // save the document
            subDocument.Save(docStream, FormatType.Docx);
            // close the document
            docStream.Dispose();
            // close the stream
            subDocument.Dispose();

        }

        public IWorksheets GetWorksheetsFromFile(string inputFileName, ExcelEngine excelEngine)
        {
            //Initialize application
            IApplication app = excelEngine.Excel;
            //Set default application version as Excel 2016
            app.DefaultVersion = ExcelVersion.Excel2016;
            //Open existing Excel workbook from the specified location
            Stream excelStream = GetStreamFromFile(inputFileName);
            IWorkbook workbook = app.Workbooks.Open(excelStream, ExcelOpenType.Automatic);

            //Access the first worksheet
            IWorksheets worksheets = workbook.Worksheets;
            return worksheets;
        }

        // get all title for worksheets
        public List<string> GetAllWorksheetsTitle(IWorksheets worksheets)
        {
            var worksheetTitles = new List<string>();
            for (int i = 0; i < worksheets.Count; i++)
            {
                worksheetTitles.Add(worksheets[i].Name);
            }

            return worksheetTitles;
        }

        // get worksheet by name 
        public IWorksheet GetWorksheetByName(IWorksheets worksheets, string worksheetName)
        {
            List<string> worksheetTitles = GetAllWorksheetsTitle(worksheets);
            // check name is exist in list
            if (worksheetTitles.Any(x => x.Trim().ToLower() == worksheetName.Trim().ToLower()))
            {
                return worksheets[worksheetName];
            }
            else
            {
                return worksheets[0];
            }
        }

        // get range from worksheet
        public IRange GetRangeFromWorksheet(IWorksheet worksheet)
        {
            //Access the used range of the Excel file
            IRange usedRange = worksheet.UsedRange;
            return usedRange;
        }
    }
}