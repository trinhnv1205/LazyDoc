using System;
using System.Collections.Generic;
using System.IO;
using LazyDoc;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.XlsIO;

// constant values
const string templateWordPath = @"/Users/trinhnv/Dotnet/LazyDoc/template.docx";
const string outputWordPath = @"/Users/trinhnv/Dotnet/LazyDoc/output.docx";
const string excelDataFile = @"/Users/trinhnv/Dotnet/LazyDoc/template.xlsx";

Dictionary<string, string> data = new Dictionary<string, string>();
Utils utils = new Utils();

// main program
Main();

// main function
void Main()
{
    utils.PrintLine("Start processing...");
    utils.PrintLine("------------------");
    // CreateDocument(templateWordPath, outputWordPath);
    ReadExcelData(excelDataFile);
    utils.PrintLine("------------------");
    utils.PrintLine("Finish processing.");
}

// create document
void CreateDocument(string inputPath, string outputPath)
{
    // check if file exists
    int checkExist = utils.CheckFileExist(inputPath);
    if (checkExist.Equals(1)) return;
    // Creates new Word document instance for word processing.
    using WordDocument document = new WordDocument();
    // Opens the input Word document.
    Stream docStream = utils.GetStreamFromFile(excelDataFile);
    // Finds all occurrences of a misspelled word and replaces with properly spelled word.
    utils.ReplaceTextDoc(document, data);
    // check exist and delete file
    utils.CheckFileExistAndDelete(outputPath);
    // Saves the resultant file in the given path.
    utils.CreateFileFromStream(docStream, outputPath);
}

void ReadExcelData(string inputFileName)
{
    //Instantiate the spreadsheet creation engine
    using ExcelEngine excelEngine = new ExcelEngine();

    IWorksheets sheets = utils.GetWorksheetsFromFile(inputFileName, excelEngine);

    IWorksheet sheet = utils.GetWorksheetByName(sheets, "Sheet1");

    IRange usedRange = utils.GetRangeFromWorksheet(sheet);

    int lastRow = usedRange.LastRow;
    int lastColumn = usedRange.LastColumn;
    
    for (int i = 1; i <= lastRow; i++)
    {
        string key = sheet[i, 1].Value;
        string value = sheet[i, 2].Value;
        data.Add(key, value);
    }

    // skip first line
    foreach (var item in data)
    {
        utils.PrintLine("Key: " + item.Key + " Value: " + item.Value);
    }
}