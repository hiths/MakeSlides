using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using ExcelManipulater;
using System.IO;

namespace TestForExcelManipulater
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = System.Environment.CurrentDirectory + "\\a.xlsx";
            Console.WriteLine(fileName);
            DataSet sheets = ExcelReader.ImportDataFromAllSheets(fileName);
            if (sheets != null)
            {
                foreach (DataTable dt in sheets.Tables)
                {
                    Console.WriteLine(dt.TableName.ToString());
                    foreach (DataRow dr in dt.Rows)
                    {
                        Console.WriteLine("line {0}:", dt.Rows.IndexOf(dr));
                        for(int i = 0; i < dt.Columns.Count; i++)
                        {
                            Console.WriteLine("{0},{1}", dr[i].ToString(), dr[i].GetType());
                            //Console.WriteLine(dr[i].ToString());
                        }
                    }       
                }
                //ExcelWriter.ExportDataToExcel(sheets, "copy.xlsx");
            }
            Console.WriteLine("Finish");
            Console.ReadKey();
        }
    }
}
