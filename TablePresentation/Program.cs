using System;
using System.Data;
using System.IO;
using ExcelManipulater;
using Newtonsoft.Json;

namespace TestForExcelManipulater
{
    class Program
    {
        //must static 
        public static void regulateData(DataTable dt)
        {
            foreach (DataRow dr in dt.Rows)
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    ((dynamic)dr[i])["color"] = Convert.ToInt32(((dynamic)dr[i])["color"]);
                    string text = ((dynamic)dr[i])["text"];
                    string format = ((dynamic)dr[i])["format"];
                    if (text.IndexOf(".") != -1 && text.IndexOf(".") == text.LastIndexOf("."))
                    {
                        
                        if(format.IndexOf("%") != -1)
                        {
                            ((dynamic)dr[i])["format"] = "0.00 % ";
                            ((dynamic)dr[i])["text"] =Math.Round(double.Parse(text), 4, MidpointRounding.AwayFromZero).ToString();
                        }
                        else
                        {
                            ((dynamic)dr[i])["text"] = Math.Round(double.Parse(text), 2, MidpointRounding.AwayFromZero).ToString();
                        }   
                    }
                }
            }
        }

        static void Main(string[] args)
        {
            string fileName = Environment.CurrentDirectory + "\\a.xlsx";
            Console.WriteLine(fileName);
            DataSet sheets = ExcelReader.ImportDataFromAllSheets(fileName);
            if (sheets != null)
            {
                foreach (DataTable dt in sheets.Tables)
                {
                    regulateData(dt);
                    string ss = JsonConvert.SerializeObject(dt, Formatting.Indented);
                    Console.WriteLine(ss);
                    File.WriteAllText(Environment.CurrentDirectory + @"WriteText.json", ss);
                }
                //ExcelWriter.ExportDataToExcel(sheets, "copy.xlsx");
            }
            Console.WriteLine("Finish");
            Console.ReadKey();
        }
    }
}
