using System;
using System.Data;
using System.IO;
using ExcelManipulater;
using Newtonsoft.Json;

namespace TestForExcelManipulater
{
    class Program
    {
        //static 
        public static void regulateData(DataTable dt)
        {
            foreach (DataRow dr in dt.Rows)
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    //string text = ((dynamic)dr[i]).text;
                    //string format = ((dynamic)dr[i]).format;
                    ((dynamic)dr[i])[1] = Convert.ToInt32(((dynamic)dr[i])[1]);
                    string text = ((dynamic)dr[i])[0];
                    string format = ((dynamic)dr[i])[2];
                    if (text.IndexOf(".") != -1 && text.IndexOf(".") == text.LastIndexOf("."))
                    {
                        
                        if(format.IndexOf("%") != -1)
                        {
                            ((dynamic)dr[i])[2] = "0.00 % ";
                            ((dynamic)dr[i])[0] =Math.Round(double.Parse(text), 4, MidpointRounding.AwayFromZero).ToString();
                        }
                        else
                        {
                            ((dynamic)dr[i])[0] = Math.Round(double.Parse(text), 2, MidpointRounding.AwayFromZero).ToString();
                        }   
                    }
                }
            }
        }

        public static DataSet ReadExcel(string excelFile)
        {
            Console.WriteLine("reading excel file named: {0}", excelFile);
            DataSet sheets = ExcelReader.ImportDataFromAllSheets(excelFile);
            string json = String.Empty;
            if (sheets != null)
            {
                foreach (DataTable dt in sheets.Tables)
                {
                    regulateData(dt);
                }
                json = JsonConvert.SerializeObject(sheets, Formatting.Indented);
                Console.WriteLine("--Data is being written to json file--");
                File.WriteAllText(excelFile + @".json", json);
                Console.WriteLine("--Write operation is complete--");
            }
            return sheets;
        }

        static void Main(string[] args)
        {
            string fileName = Environment.CurrentDirectory + "\\a.xlsx";
            ReadExcel(fileName);
            Console.WriteLine("Finish");
            Console.ReadKey();
        }
    }
}
