using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace FetchCell
{
    public class Cell
    {
        static  void Main()
        {
            object Nothing = System.Reflection.Missing.Value;
            var xlApp = new Excel.Application();
            xlApp.Visible = true;
            Excel.Workbook workBook = xlApp.Workbooks.Add(Nothing);
            Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Sheets[1];
            workSheet.Name = "sheet0n3";
            workSheet.Cells[1, 1] = "FileName";
            workSheet.Cells[1, 2] = "FindString";
            workSheet.Cells[1, 3] = "ReplaceString";

            workSheet.SaveAs("1.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);
            workBook.Close(false, Type.Missing, Type.Missing);
            xlApp.Quit();
        }
    }

}
