using System;
using System.Data;
using System.IO;
using ExcelManipulater;
using PowerPointOperator;
using Newtonsoft.Json;
using System.Linq;
using System.Collections.Generic;

namespace MakeSlidesFromExcel
{
    class Program
    {
        private Dictionary<string, int> games;

        public static Dictionary<string, int> getCustomization(string filePath)
        {

            //return getCustomization;
        }

        //must static 
        public static void regulateData(DataTable dt, int width = 0)
        {
            foreach (DataRow dr in dt.Rows)
            {
                for (int i = 0; i < width; i++)
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
            Console.WriteLine("reading excel file : {0}", excelFile);
            DataSet sheets = ExcelReader.ImportDataFromAllSheets(excelFile);
            string json = String.Empty;
            if (sheets != null)
            {
                if (whiteList != null)
                {
                    for(int i = sheets.Tables.Count-1; i >= 0; i --)
                    {
                        string tableName = sheets.Tables[i].TableName;
                        if (!whiteList.Contains(tableName))
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

        public static DataSet makeStructure(DataSet newSheets, DataSet exsitedStructure = null)
        {
            DataSet newStructure = new DataSet();
            if(exsitedStructure != null)
            {
                Dictionary<string, string[]> slidesIndex = miniStructure(exsitedStructure);
                for (int i = 0; i < newSheets.Tables.Count; i++)
                {
                    string game = exsitedStructure.Tables[i].TableName.Split(new char[1] { '-' })[0];
                    if (slidesIndex.Keys.Contains(game))
                    {
                        for (int j = 1; j < newSheets.Tables[i].Rows.Count; j ++ )
                        {
                            string mat = ((dynamic)newSheets.Tables[i].Rows[j])[0]["text"];
                            DataRow dr = newSheets.Tables[i].Rows[j];
                            if (slidesIndex[game].Contains(mat))
                            {

                            }
                        }
                    }

                }
            }
            else {
                for (int i = 0; i < newSheets.Tables.Count; i++)
                {
                    DataTable dt = newSheets.Tables[i]; 
                    for (int j = 1; j < dt.Rows.Count; j++)
                    {
                        DataTable newTable = new DataTable(); // in or out of for sentance ?
                        newTable.TableName = dt.TableName.ToString() + "-" + ((dynamic)dt.Rows[j])["text"];
                        newTable.Rows.Add(dt.Rows[0]);
                        newTable.Rows.Add(dt.Rows[j]);
                        newStructure.Tables.Add(newTable);
                    }
                }
            }
            return newStructure;
        }

        public static Dictionary<string, string[]>  miniStructure(DataSet struture)
        {
            Dictionary<string, string[]> miniStructure = new Dictionary<string, string[]>();

            return miniStructure;
        }

        public static string[] getSlidesIndex(DataSet structure)
        {
            int k = structure.Tables.Count;
            /string[] sildesIndex = new string[]();
            /return slidesIndex;
        }

        static void Main(string[] args)
        {
            string excelName = Environment.CurrentDirectory + "\\a.xlsx";
            string[] games = new string[2] { "大皇帝", "少年三国志"};
            DataSet sheets = ReadExcel(excelName,games);
            /*
            string pptName = Environment.CurrentDirectory + "\\test.pptx";
            SlidesEditer.openPPT(pptName);
            */
            Console.WriteLine("Finish");
            Console.ReadKey();
        }
    }
}
