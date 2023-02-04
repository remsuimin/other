using System;
using Microsoft.Office.Interop.Excel;

class Program
{
    static void Main(string[] args)
    {
        Application excel = new Application();
        Workbook workbook = excel.Workbooks.Open("test.xlsx");

        workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, "test.pdf");

        workbook.Close();
        excel.Quit();
    }
}
