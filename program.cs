using System;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Linq;
using System.Collections;
using System.Text.RegularExpressions;
using Bytescout.Spreadsheet;
using System.Text;

namespace ConsoleApp3
{
    class Program
    {
        private static string[] ReadAllFilesInDirectory(string topLevelDirectory, string[] arr)
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
                arrFunc = arrFunc.Concat(ReadAllFilesInDirectory(subDirectory, arrFunc)).Distinct().ToArray();//recursion
            }

            arrFunc = arrFunc.Concat(filesInDirectory).ToArray();
            return arrFunc;
        }
        private static string GetPlainTextFromHtml(string htmlString)
        {
            string htmlTagPattern = "<.*?>";
            var regexCss = new Regex("(\\<script(.+?)\\</script\\>)|(\\<style(.+?)\\</style\\>)", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            htmlString = regexCss.Replace(htmlString, string.Empty);
            htmlString = Regex.Replace(htmlString, htmlTagPattern, string.Empty);
            htmlString = Regex.Replace(htmlString, @"^\s+$[\r\n]*", "", RegexOptions.Multiline);
            htmlString = htmlString.Replace("&nbsp;", string.Empty);

            return htmlString;
        }
        private static string[][] IterateFiles(IEnumerable<string> files)
        {
            string[][] termsList = new string[0][];

            foreach (var file in files)
            {
                try
                {
                    StringBuilder sb = new StringBuilder();
                    string[] lines = File.ReadAllLines(file);
                    foreach (var line in lines)
                    {
                        string convert = GetPlainTextFromHtml(line);
                        //handle check line on string;
                        if (!String.IsNullOrEmpty(convert.Trim()))
                        {
                            if (!convert.Trim().StartsWith("{") && !convert.Trim().EndsWith("}"))
                            {
                                string[] fileExtention = { file, convert };
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
            string[] arr = new string[0];
            string[] tes = ReadAllFilesInDirectory(@"C:\Users\ADMIN\Desktop\Readfile\test", arr);
            string[][] readFile = new string[0][];
            readFile = IterateFiles(tes);
            for (int i = 2; i < readFile.Length; i++)
            {

                Sheet.Cell(Convert.ToString("A" + i)).Value = readFile[i][0].Trim();
                Sheet.Cell(Convert.ToString("B" + i)).Value = readFile[i][1].Trim();
            }
            document.SaveAs("Output.xls");
            document.Close();
            Process.Start("Output.xls");
        }

    }
}
