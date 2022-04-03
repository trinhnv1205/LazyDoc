using System;
using System.Collections.Generic;
using System.IO;
using Syncfusion.DocIO.DLS;

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
                document.Replace(item.Key.ToString(), item.Value.ToString(), true, true);
            }
        }
    }
}