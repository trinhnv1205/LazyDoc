using System.Collections.Generic;
using System.IO;
using LazyDoc;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

// constant values
const string input = @"/Users/trinhnv/Dotnet/LazyDoc/template.docx";
const string output = @"/Users/trinhnv/Dotnet/LazyDoc/output.docx";

Dictionary<string, string> data = new Dictionary<string, string>()
{
    {"4", "TrinhNV"},
    {"5", "20"},
    {"1", "Ha Noi"}
};
Utils utils = new Utils();

// main program
Main();

// main function
void Main()
{
    utils.PrintLine("Start processing...");
    CreateDocument(input, output);
    utils.PrintLine("Finish processing.");
}

// create document
void CreateDocument(string inputPath, string outputPath)
{
    // check if file exists
    int checkExist = utils.CheckFileExist(inputPath);
    if (checkExist.Equals(1))
    {
        return;
    }

    // Creates new Word document instance for word processing.
    using WordDocument document = new WordDocument();
    // Opens the input Word document.
    Stream docStream = File.OpenRead(Path.GetFullPath(inputPath));
    document.Open(docStream, FormatType.Docx);
    docStream.Dispose();

    // Finds all occurrences of a misspelled word and replaces with properly spelled word.
    utils.ReplaceTextDoc(document, data);

    // check exist and delete file
    utils.CheckFileExistAndDelete(outputPath);
    // Saves the resultant file in the given path.
    docStream = File.Create(Path.GetFullPath(outputPath));
    document.Save(docStream, FormatType.Docx);
    docStream.Dispose();
}