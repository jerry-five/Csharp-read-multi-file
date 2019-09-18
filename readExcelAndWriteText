using System;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
class Program
{
    public static void ReadSample()
    {
        Excel.Application excelApp = new Excel.Application();
        if (excelApp != null)
        {
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(@"C:\Users\ADMIN\Desktop\Malay\malay.xlsx", 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets[1];

            Excel.Range excelRange = excelWorksheet.UsedRange;
            int rowCount = excelRange.Rows.Count;
            int colCount = excelRange.Columns.Count;

            for (int i = 1; i <= rowCount; i++)
            {
                if (i > 1 && (excelWorksheet.Cells[i, 4]).Value2 != null)
                {
                    Excel.Range line = (excelWorksheet.Cells[i, 2]);

                    Excel.Range path = (excelWorksheet.Cells[i, 5]);
                    Excel.Range newText = (excelWorksheet.Cells[i, 4]);
                    Excel.Range englishText = (excelWorksheet.Cells[i, 3]);
                    if (line.Value != null && path.Value.Length > 0 && newText.Value.ToString().Length > 0 && englishText.Value.Length > 0)
                    {
                        int valueLine = Convert.ToInt32(line.Value);
                        string valuePath = path.Value.ToString();
                        string ValueText = newText.Value.ToString();
                        string textCompare = englishText.Value.ToString();
                        lineChanger(ValueText, valuePath, valueLine, textCompare);
                    }

                }
            }

            excelWorkbook.Close();
            excelApp.Quit();
        }
    }
    static void lineChanger(string newText, string fileName, int line_to_edit, string textCompare)
    {
        string[] arrLine = File.ReadAllLines(fileName);
        //Console.WriteLine("textCompare " + textCompare);
        //Console.WriteLine("line number " + line_to_edit);
        //Console.WriteLine("arrLine[line_to_edit] " + arrLine[line_to_edit - 1]);
        string upText = arrLine[line_to_edit - 1].Replace(textCompare, newText);
        Console.WriteLine("fileNam " + fileName);
        Console.WriteLine("replace Text " + upText);
        Console.WriteLine("==============================\n");
        arrLine[line_to_edit - 1] = upText;
        //  Console.WriteLine("arrLine[line_to_edit -1]" + arrLine[line_to_edit - 1]);
        File.WriteAllLines(fileName, arrLine);
    }
    static void Main()
    {
        ReadSample();
        // string tes = @"C:\Users\ADMIN\Desktop\Test.xml";
        //string[] lines = System.IO.File.ReadAllLines(@"C:\Users\ADMIN\Desktop\Test.xml");
        //for (int i = 0; i <= lines.Length; i++)
        //{
        //    if (i == 2)
        //    {
        //        string addXml = lines[3].Replace("Mr.,Ms./Mrs.,Child", "tai");
        //        lineChanger(addXml, tes, 4);
        //    }
        //}
        Console.WriteLine("Press any key to exit.");
        System.Console.ReadKey();
    }
}
