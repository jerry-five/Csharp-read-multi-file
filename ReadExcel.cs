using System;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
class Program
{
    public static void ReadSample()
    {
        Excel.Application excelApp = new Excel.Application();
        if (excelApp != null)
        {
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(@"C:\Users\ADMIN\Downloads\translate.xlsx", 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets[1];

            Excel.Range excelRange = excelWorksheet.UsedRange;
            int rowCount = excelRange.Rows.Count;
            int colCount = excelRange.Columns.Count;

            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j < colCount; j++)
                {
                    Excel.Range range = (excelWorksheet.Cells[i, j]);
                    string cellValue = range.Value.ToString();
                    Console.WriteLine(cellValue);
                    //do anything
                }
            }

            excelWorkbook.Close();
            excelApp.Quit();
        }
    }
    static void Main()
    {
        ReadSample();
    }
}
