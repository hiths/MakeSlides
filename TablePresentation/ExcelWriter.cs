﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ExcelManipulater
{
    public class ExcelWriter
    {
        public static Boolean ExportDataToExcel(DataSet dataSheets, string fileName)
        {
            Excel.Application xlApp = null;
            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet = null;
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            //xlApp.Visible = true;
            xlApp.DisplayAlerts = false;
            int sheetNum = dataSheets.Tables.Count;
            int i = 0;
            int j = 0;

            try
            {
                for (int sheetIndex = 1; sheetIndex <= sheetNum; sheetIndex++)
                {
                    Console.WriteLine("Writing Sheet" + sheetIndex);
                    if (xlWorkBook.Sheets.Count < sheetIndex)
                    {
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets.Add(misValue, xlWorkBook.Sheets[sheetIndex - 1], misValue, misValue);
                    }
                    else
                    {
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(sheetIndex);
                    }
                    DataTable dataTable = dataSheets.Tables[sheetIndex - 1];
                    xlWorkSheet.Name = dataTable.TableName;
                    for (i = 0; i < dataTable.Rows.Count; i++)
                    {
                        DataRow dr = dataTable.Rows[i];
                        for (j = 0; j < dataTable.Columns.Count; j++)
                        {
                            xlWorkSheet.Cells[i + 1, j + 1] = ((dynamic)dr.ItemArray[j])["text"].ToString();
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                return false;
            }
            finally
            {
                xlWorkBook.SaveAs(fileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, fileName, misValue);
                xlApp.Quit();

                ReleaseObject(xlWorkSheet);
                ReleaseObject(xlWorkBook);
                ReleaseObject(xlApp);
            }

            return true;
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

        [DllImport("User32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int ProcessId);
        private static void KillExcel(Excel.Application theApp)
        {
            int id = 0;
            IntPtr intptr = new IntPtr(theApp.Hwnd);
            System.Diagnostics.Process p = null;
            try
            {
                GetWindowThreadProcessId(intptr, out id);
                p = System.Diagnostics.Process.GetProcessById(id);
                if (p != null)
                {
                    p.Kill();
                    p.Dispose();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
    }
}