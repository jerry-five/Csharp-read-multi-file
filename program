using System;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections;

namespace ConsoleApp3
{
    class Program
    {
        public static ArrayList patchExcel = new ArrayList(); ///public variable declaration

        public static int ReadAllFilesStartingFromDirectory(string topLevelDirectory)
        {
            const string searchPattern = "*.xml";
            const string searchPattern1 = "*.asp";
            var subDirectories = Directory.EnumerateDirectories(topLevelDirectory);
            Console.WriteLine("topLevelDirectory"+ topLevelDirectory);
            Console.WriteLine("======================================");
            var filesInDirectory1 = Directory.EnumerateFiles(topLevelDirectory, searchPattern);
            var filesInDirectory2 = Directory.EnumerateFiles(topLevelDirectory, searchPattern1);
            var filesInDirectory = filesInDirectory1.Concat(filesInDirectory2).ToArray();
            //Console.WriteLine(filesInDirectory);
            foreach (var subDirectory in subDirectories)
            {
                Console.WriteLine("subDirectories"+ subDirectories);
                Console.WriteLine("subDirectory"+ subDirectory);
                Console.WriteLine("======================================");
                ReadAllFilesStartingFromDirectory(subDirectory);//recursion;
            }

            foreach (var file in filesInDirectory)
            {
               // Console.WriteLine(filesInDirectory.Length);
                // Console.WriteLine("{0}", Path.Combine(topLevelDirectory, file));
                try
                {
                    Console.WriteLine("file"+ Path.Combine(topLevelDirectory, file));
                    string[] lines = File.ReadAllLines(file);
                    foreach (var line in lines)
                    {
                        //write line on excel;
                    }
                }
                catch (IOException ex)
                {
                    System.Threading.Thread.Sleep(100);
                    //Handle File may be in use...       
                    throw ex;
                }
            }
            return 1;
        }
        static void Main(string[] args)
        {
            var excel = new Excel.Application();

            var workBooks = excel.Workbooks;
            var workBook = workBooks.Add();
            var workSheet = (Excel.Worksheet)excel.ActiveSheet;

            workSheet.Cells[1, "A"] = "Path";
            workSheet.Cells[1, "B"] = "English";
            DateTime foo = DateTime.UtcNow;
            long unixTime = ((DateTimeOffset)foo).ToUnixTimeSeconds();


            string a = Convert.ToString(unixTime + ".xls");
            workBook.SaveAs(Directory.GetCurrentDirectory() + "\\" + a, Excel.XlFileFormat.xlOpenXMLWorkbook);
            workBook.Close();
            string[] fileArray = Directory.GetFiles(@"C:\Users\Donald-Trump\Desktop\Readfile", "*.asp");
           // Console.WriteLine(fileArray);
            foreach (string fileName in fileArray)  
            {
                Console.WriteLine(fileName);
            }
            ReadAllFilesStartingFromDirectory("C:\\Users\\Donald-Trump\\Desktop\\Readfile");
          //  Console.WriteLine(patchExcel);
        }
    }
}
