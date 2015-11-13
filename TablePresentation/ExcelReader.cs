using System;
using System.IO;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelManipulater
{
    public class ExcelReader
    {
        private static void Initialize(string fileName, out Excel.Application xlApp, out Excel.Workbook xlWorkBook)
        {
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(fileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
        }

        public static DataSet ImportDataFromAllSheets(string fileName)
        {
            if (!File.Exists(fileName))
            {
                return null;
            }

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Initialize(fileName, out xlApp, out xlWorkBook);

            int workSheetNum = xlWorkBook.Worksheets.Count;
            int sheetCount=0;
            DataSet sheets = new DataSet();
            Console.WriteLine("WorkSheet number:" + workSheetNum);
            try
            {
                for (sheetCount = 1; sheetCount <= workSheetNum; sheetCount++)
                {
                    Console.WriteLine("Reading sheet" + sheetCount);
                    DataTable sheetData = ExtractDataFromSingleSheet(xlWorkBook, sheetCount);
                    sheets.Tables.Add(sheetData);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            finally
            {
                Dispose(fileName, ref xlApp, ref xlWorkBook);
            }

            return sheets;
        }

        public static DataTable ImprotDataFromSingleSheet(string fileName, int sheetIndex)
        {
            if(!File.Exists(fileName))
            {
                return null;
            }

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Initialize(fileName, out xlApp, out xlWorkBook);

            DataTable sheetData = null;

            try
            {
                sheetData = ExtractDataFromSingleSheet(xlWorkBook, sheetIndex);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            finally
            {
                Dispose(fileName, ref xlApp, ref xlWorkBook);
            }

            return sheetData;
        }

        private static DataTable ExtractDataFromSingleSheet(Excel.Workbook xlWorkBook, int sheetCount)
        {
            int rowCount = 0;
            int columnCount = 0;
            DataTable sheetData = new DataTable();

            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(sheetCount);
            Excel.Range range = xlWorkSheet.UsedRange;
            sheetData.TableName = xlWorkSheet.Name;

            for (int i = 0; i < range.Columns.Count; i++)
            {
                sheetData.Columns.Add(new DataColumn());
            }

            try
            {
                for (rowCount = 1; rowCount <= range.Rows.Count; rowCount++)
                {
                    DataRow row = sheetData.NewRow();
                    for (columnCount = 1; columnCount <= range.Columns.Count; columnCount++)
                    {
                        object cellText = (range.Cells[rowCount, columnCount] as Excel.Range).Value2;
                        double textColor = 0;
                        if (cellText == null)
                        {
                            cellText = String.Empty;
                        }
                        else
                        {
                            cellText = (range.Cells[rowCount, columnCount] as Excel.Range).Value2.ToString();
                            textColor = (range.Cells[rowCount, columnCount] as Excel.Range).Font.Color;
                        }
                        var cell = new[] {cellText, textColor };
                        row[columnCount - 1] = cell;
                        //row[columnCount - 1] = cellText != null ? (range.Cells[rowCount, columnCount] as Excel.Range).Value2.ToString() : String.Empty;
                    }
                    sheetData.Rows.Add(row);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            finally
            {
                ReleaseObject(xlWorkSheet);
            }

            return sheetData;
        }

        private static void Dispose(string fileName, ref Excel.Application xlApp, ref Excel.Workbook xlWorkBook)
        {
            xlWorkBook.Close(false, fileName, null);
            xlApp.Quit();

            ReleaseObject(xlWorkBook);
            ReleaseObject(xlApp);
        }

        private static void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
