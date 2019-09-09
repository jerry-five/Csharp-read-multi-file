using System;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections;
using Bytescout.Spreadsheet;

namespace ConsoleApp3
{
    class Program
    {
        private static string[] ReadAllFilesStartingFromDirectory(string topLevelDirectory, string[] arr )
        {
            string[] arrFunc = arr;
            const string searchPattern = "*.xml";
            const string searchPattern1 = "*.asp";
            var subDirectories = Directory.EnumerateDirectories(topLevelDirectory);
            var filesInDirectory1 = Directory.EnumerateFiles(topLevelDirectory, searchPattern);
            var filesInDirectory2 = Directory.EnumerateFiles(topLevelDirectory, searchPattern1);
            var filesInDirectory = filesInDirectory1.Concat(filesInDirectory2).ToArray();
            foreach (var subDirectory in subDirectories)
            {
                ReadAllFilesStartingFromDirectory(subDirectory, arrFunc);//recursion
                arrFunc = arrFunc.Concat(IterateFiles(filesInDirectory, topLevelDirectory)).ToArray();
            }
               // arrFunc= arrFunc.Concat(IterateFiles(filesInDirectory, topLevelDirectory)).ToArray();
            return arrFunc;
        }

        private static string[] IterateFiles(IEnumerable<string> files, string directory)
        {
            string[] termsList = new string[0];

            foreach (var file in files)
            {
               // Console.WriteLine("{0}", Path.Combine(directory, file));//for verification
                try
                {
                    string[] lines = File.ReadAllLines(file);
                    foreach (var line in lines)
                    {
                        termsList = termsList.Concat(new string[] { line }).ToArray();

                        //Console.WriteLine(line); 
                    }
                }
                catch (IOException ex)
                {
                    throw ex;
                    //Handle File may be in use...                    
                }
            }
            return termsList;
        }
        static void Main(string[] args)
        {
            Spreadsheet document = new Spreadsheet();

            // add new worksheet
            Worksheet Sheet = document.Workbook.Worksheets.Add("FormulaDemo");

            // headers to indicate purpose of the column
            Sheet.Cell("A1").Value = "Formula (as text)";
            // set A column width
            Sheet.Columns[0].Width = 250;

            Sheet.Cell("B1").Value = "Formula (calculated)";
            // set B column width
            Sheet.Columns[1].Width = 250;

            string[] termsList = new string[5];
           string[] tes = ReadAllFilesStartingFromDirectory(@"C:\Users\ADMIN\Desktop\Readfile", termsList);
            for (int i = 2; i < tes.Length; i++)
            {
                Sheet.Cell(Convert.ToString("A" + i)).Value = i;
                Sheet.Cell(Convert.ToString("A" + i)).Value = tes[i];
            }
            document.SaveAs("Output.xls");

            // Close Spreadsheet
            document.Close();

            // open generated XLS document in default program
            Process.Start("Output.xls");
        }


    }
}
