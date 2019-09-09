using System;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Linq;
using System.Collections;
using Bytescout.Spreadsheet;
using System.Text.RegularExpressions;

namespace ConsoleApp3
{
    class Program
    {
        private static string[][] ReadAllFilesInDirectory(string topLevelDirectory, string[][] arr)
        {
            string[][] arrFunc = arr;
            const string searchPattern = "*.xml";
            const string searchPattern1 = "*.asp";
            var subDirectories = Directory.EnumerateDirectories(topLevelDirectory);
            var filesInDirectory1 = Directory.EnumerateFiles(topLevelDirectory, searchPattern);
            var filesInDirectory2 = Directory.EnumerateFiles(topLevelDirectory, searchPattern1);
            var filesInDirectory = filesInDirectory1.Concat(filesInDirectory2).ToArray();
            foreach (var subDirectory in subDirectories)
            {
                arrFunc = arrFunc.Concat(ReadAllFilesInDirectory(subDirectory, arrFunc)).ToArray();//recursion
            }
            arrFunc = arrFunc.Concat(IterateFiles(filesInDirectory, topLevelDirectory)).ToArray();
            return arrFunc;
        }

        private static string[][] IterateFiles(IEnumerable<string> files, string directory)
        {
            string[][] termsList = new string[0][];

            foreach (var file in files)
            {
                try
                {
                    string[] lines = File.ReadAllLines(file);
                    foreach (var line in lines)
                    {
                        if(!String.IsNullOrEmpty(line.Trim()))
                        {

                            String output = Regex.Replace(line, @"\>([^\[\]]+)\<", "");
                            if(output != null)
                            {
                                string[] fileExtention = { file, output };
                                termsList = termsList.Concat(new string[][] { fileExtention }).ToArray();
                            }
                        }
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
            Worksheet Sheet = document.Workbook.Worksheets.Add("sheet1");
            Sheet.Cell("A1").Value = "Path";
            Sheet.Columns[0].Width = 250;
            Sheet.Cell("B1").Value = "English";
            Sheet.Columns[1].Width = 250;
            string[][] termsList = new string[0][];
            string[][] tes = ReadAllFilesInDirectory(@"C:\Users\ADMIN\Desktop\Readfile\test", termsList);
            for (int i = 2; i < tes.Length; i++)
            {
                Sheet.Cell(Convert.ToString("A" + i)).Value = tes[i][0];
                Sheet.Cell(Convert.ToString("B" + i)).Value = tes[i][1];
            }
            document.SaveAs("Output.xls");

            document.Close();
            Process.Start("Output.xls");
        }

    }
}
