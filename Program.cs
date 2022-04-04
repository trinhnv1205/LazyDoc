using System.Collections.Generic;
using LazyDoc;
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
    ReadExcelData(excelDataFile);
    CreateDocument(templateWordPath, outputWordPath);
    utils.PrintLine("------------------");
    utils.PrintLine("Finish processing.");
}

// create document
void CreateDocument(string inputPath, string outputPath)
{
    // check if file exists
    if (utils.CheckFileExist(inputPath).Equals(1)) return;
    // check exist and delete file
    utils.CheckFileExistAndDelete(outputPath);
    // Saves the resultant file in the given path.
    utils.CreateFileFromStream(data, templateWordPath, outputPath);
    //Saves the resultant file in the given path.
}

// read excel data
void ReadExcelData(string inputFileName)
{
    //Instantiate the spreadsheet creation engine
    using ExcelEngine excelEngine = new ExcelEngine();
    IWorksheets sheets = utils.GetWorksheetsFromFile(inputFileName, excelEngine);
    IWorksheet sheet = utils.GetWorksheetByName(sheets, "Sheet1");
    IRange usedRange = utils.GetRangeFromWorksheet(sheet);

    //Read the data from the spreadsheet
    int lastRow = usedRange.LastRow;
    // for each row to get data from excel file and add to dictionary
    for (int i = 1; i <= lastRow; i++)
    {
        string key = sheet[i, 1].Value.Trim();
        string value = sheet[i, 2].Value.Trim();
        // check if key is empty or value is empty -> skip
        if(string.IsNullOrEmpty(key) || string.IsNullOrEmpty(value)) continue;
        // check if key is existed in dictionary
        if (utils.CheckKeyExist(data, key))
        {
            // update value
            data[key] = value;
        }
        else{
            // add new key and value
            data.Add(key, value);
        }
    }
}