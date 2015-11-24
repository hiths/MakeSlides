using System;
using System.Data;
using System.IO;
using ExcelManipulater;
using PowerPointOperator;
using Newtonsoft.Json;
using System.Linq;

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

        public static DataSet ReadExcel(string excelFile, string[] whiteList = null)
        {
            Console.WriteLine("reading excel file named: {0}", excelFile);
            DataSet sheets = ExcelReader.ImportDataFromAllSheets(excelFile);
            string json = String.Empty;
            if (sheets != null)
            {
                if (whiteList != null)
                {
                    //foreach (DataTable dt in sheets.Tables)
                    for(int i = sheets.Tables.Count-1; i >= 0; i --)
                    {
                        if (!whiteList.Contains(sheets.Tables[i].TableName))
                        {
                            sheets.Tables.Remove(sheets.Tables[i]);
                        }
                        else
                        {
                            regulateData(sheets.Tables[i]);
                        }
                    }
                }
                else
                {
                    foreach (DataTable dt in sheets.Tables)
                    {
                        regulateData(dt);
                    }
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
            //string excelName = Environment.CurrentDirectory + "\\a.xlsx";
            //string[] games = new string[2] { "大皇帝", "少年三国志"};
            //DataSet sheets = ReadExcel(excelName,games);
            string pptName = Environment.CurrentDirectory + "\\test.pptx";
            SlidesEditer.openPPT(pptName);
            Console.WriteLine("Finish");
            Console.ReadKey();
        }
    }
}
