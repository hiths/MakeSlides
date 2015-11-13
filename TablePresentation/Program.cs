using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace FetchCell
{
    public class Cell
    {
        static void Main()
        {
            object Nothing = System.Reflection.Missing.Value;
            var xlApp = new Excel.Application();
            xlApp.Visible = true;
            Excel.Workbook workBook = xlApp.Workbooks.Add(Nothing);
            Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Sheets[1];
            Excel.Range range = workSheet.UsedRange;
            workSheet.Name = "sheet";
            workSheet.Cells[1, 1] = "FileName";
            workSheet.Cells[1, 2] = "FindString";
            workSheet.Cells[1, 3] = "ReplaceString";
            Excel.Range thisCell = range.Cells[1,1];
            Object value = thisCell.Value;
            //fontcolor
            double fontColor = thisCell.Font.Color;
            //background color
            double bgColor = thisCell.Interior.Color;
            //cell note
            //string note = "sb";
            //thisCell.Comment.Text = note;
            var thisArray = new []{ value, fontColor, bgColor };
            foreach(var i in thisArray)
            {
                Console.WriteLine(i);
            };
            Console.ReadKey();
            Console.WriteLine(thisCell.Comment.Text());

            //workSheet.SaveAs("1.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);
            //workBook.Close(false, Type.Missing, Type.Missing);
            //xlApp.Quit();
        }
    }

}
