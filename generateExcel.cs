using System;
using Microsoft.Office.Interop.Excel;

class Program
{
    static void Main(string[] args)
    {
        Application excel = new Application();
        Workbook workbook = excel.Workbooks.Add();
        Worksheet worksheet = workbook.Worksheets.get_Item(1);

        for (int row = 1; row <= 3; row++)
        {
            for (int col = 1; col <= 3; col++)
            {
                worksheet.Cells[row, col].Value = (row - 1) * 3 + col;
            }
        }

        workbook.SaveAs("test.xlsx");
        workbook.Close();
        excel.Quit();
    }
}
