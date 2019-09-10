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

        private static string[][] IterateFiles(IEnumerable<string> files, int stringLength)
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
                        string convert = GetPlainTextFromHtml(line).Trim();
                        //handle check line on string;
                        if (!String.IsNullOrEmpty(convert))
                        {
                            if (
                                   !convert.StartsWith("{")
                                && !convert.EndsWith("}")
                                && (!convert.Contains("=") || !convert.Contains("."))
                                && !convert.Contains("<%")
                                && !convert.Contains("%>")
                                && !convert.Contains("NOTE: ")
                                && !convert.Contains("-->")
                                && !convert.Contains("&#")
                                && !convert.Contains("javascript")
                                && !convert.Contains("/*")
                                && !convert.Contains("*/")
                                && !convert.Contains(".jpg")
                                && !convert.Contains("vvSelect(")
                                && !convert.Contains("/script")
                                && !convert.Contains("'script'")
                                && !convert.Contains("http")
                                )
                            {
                                string[] fileExtention = { file.Remove(0, stringLength + 1), convert };
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
            string[] arr = new string[0];
            string[][] readFile = new string[0][];
            string pathUrl = @"C:\Users\ADMIN\Desktop\Readfile";
            Spreadsheet document = new Spreadsheet();
            Worksheet Sheet = document.Workbook.Worksheets.Add("sheet1");

            Sheet.Cell("A1").Value = "Path";
            Sheet.Columns[0].Width = 250;
            Sheet.Cell("B1").Value = "English";
            Sheet.Columns[1].Width = 250;

            string[] tes = ReadAllFilesInDirectory(pathUrl, arr);
            Console.WriteLine(pathUrl.Length);
            readFile = IterateFiles(tes, pathUrl.Length);
            int rowIndex = 2;

            for (int i = 0; i < readFile.Length; i++)
            {

                Sheet.Cell(Convert.ToString("A" + rowIndex)).Value = readFile[i][0];
                Sheet.Cell(Convert.ToString("B" + rowIndex)).Value = readFile[i][1];
                rowIndex++;
            }

            document.SaveAs("Output.xls");
            document.Close();
            Process.Start("Output.xls");
        }

    }
}
