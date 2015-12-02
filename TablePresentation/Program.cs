using System;
using System.Data;
using System.IO;
using ExcelManipulater;
using PowerPointOperator;
using Newtonsoft.Json;
using System.Linq;
using System.Collections.Generic;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Newtonsoft.Json.Linq;

namespace MakeSlidesFromExcel
{
    class Program
    {
        private static Dictionary<string, int[]> gameConfig = new Dictionary<string, int[]>();
        private static int gamesCount = 0;
        private static List<string> gameList = new List<string>();
        private static string configFile = Environment.CurrentDirectory + "\\GamesInfo.json";
        private static string structureFile = Environment.CurrentDirectory + "\\SlidesMap.json";
        //private static string configFileFolder = "Customized";
        //private static string backupFolder = "Backup";
        //private static string outputFolder = "OutPut";
        private static string projectFolder = Environment.CurrentDirectory + "\\Project";

        public static void initialize()
        {
            if (File.Exists(configFile))
            {
                gameConfig = getConfigFile(configFile);
                gamesCount = gameConfig.Keys.Count;
                foreach(string game in gameConfig.Keys)
                {
                    Console.WriteLine(game);
                    gameList.Add(game);
                }
            }
            else
            {
                File.Create(configFile);
                Console.WriteLine("Build a config file before the initialization of a project.");
                Console.WriteLine("Press any key to exit.");
                Console.ReadKey();
                Environment.Exit(0);
            }

            if (!File.Exists(projectFolder))
            {
                Directory.CreateDirectory(projectFolder);
            }

        }

        public static Dictionary<string, int[]> getConfigFile(string configFilePath)
        {
            Dictionary<string, int[]> customization = new Dictionary<string, int[]>();
            string rawJson = File.ReadAllText(@"GamesInfo.json");
            Console.WriteLine(rawJson);
            customization = JsonConvert.DeserializeObject<Dictionary<string, int[]>>(rawJson);
            return customization;
        }

        public static void regulateData(DataTable dt, int width = 8)
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
                            ((dynamic)dr[i])["text"] =(Math.Round(double.Parse(text), 4, MidpointRounding.AwayFromZero)*100).ToString() + "%";

                        }
                        else
                        {
                            ((dynamic)dr[i])["text"] = Math.Round(double.Parse(text), 2, MidpointRounding.AwayFromZero).ToString();
                        }   
                    }
                    dr[i] = new Dictionary<string, object> { { "text", ((dynamic)dr[i])["text"]}, { "color", ((dynamic)dr[i])["color"]}};
                }
            }
        }

        public static DataSet ReadExcel(string excelFile, Dictionary<string, int[]> gameConfig = null)
        {
            if (!File.Exists(excelFile))
            {
                Console.WriteLine("Can not find specified excel file.");
                Console.ReadKey();
                Environment.Exit(0);
            }
            Console.WriteLine("reading excel file : {0}", excelFile);
            DataSet sheets = ExcelReader.ImportDataFromAllSheets(excelFile);
            string json = String.Empty;
            if (sheets != null)
            {
                if (gameConfig != null)
                {
                    for(int i = sheets.Tables.Count-1; i >= 0; i --)
                    {
                        string tableName = sheets.Tables[i].TableName;
                        if (!gameConfig.Keys.Contains(tableName))
                        {
                            sheets.Tables.Remove(sheets.Tables[i]);
                        }
                        else
                        {
                            int width = ((dynamic)gameConfig[tableName])[2];
                            regulateData(sheets.Tables[i], width+1);
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

        public static DataSet insertTableToSet(DataSet ds, DataTable dt, int index)
        {
            if(index < 0 | index >= ds.Tables.Count)
            {
                ds.Tables.Add(dt);
                return ds;
            }
            else
            {
                DataSet newDataSet = new DataSet();
                for(int i = 0; i < index; i++)
                {
                    newDataSet.Tables.Add(ds.Tables[i].Copy());
                }
                newDataSet.Tables.Add(dt);
                for(int i = index; i < ds.Tables.Count; i++)
                {
                    newDataSet.Tables.Add(ds.Tables[i].Copy());
                }
                return newDataSet;
            }
        }

        public static DataSet makeStructure(PowerPoint.Presentation pptPrest, DataSet newSheets, DataSet structure = null)
        {
            int pageIndexOfThisRow = 0;
            if (structure != null)
            {
                List<object> slidesIndex = getSlidesIndex(structure);        
                for (int i = 0; i < newSheets.Tables.Count; i++)
                {
                    string game = structure.Tables[i].TableName.Split(new char[1] { '-' })[0];
                    if (((dynamic)slidesIndex[0]).Contains(game))
                    {
                        for (int j = 1; j < newSheets.Tables[i].Rows.Count; j ++ )
                        {
                            string column_0 = ((dynamic)newSheets.Tables[i].Rows[j])[0]["text"];
                            DataRow dr = newSheets.Tables[i].Rows[j];
                            if (((dynamic)slidesIndex[1]).Contains(column_0) && ((dynamic)slidesIndex[0])[((dynamic)slidesIndex[1]).IndexOf(column_0)] == game)
                            {
                                int index = ((dynamic)slidesIndex[1]).IndexOf(column_0);
                                structure.Tables[index].Rows.Add(dr.ItemArray);
                            }
                            else
                            {
                                int index = ((dynamic)slidesIndex[0]).LastIndexOf(game);
                                DataTable newTable = new DataTable(game + "-" + column_0);
                                newTable.Rows.Add(newSheets.Tables[i].Rows[0].ItemArray);
                                newTable.Rows.Add(dr.ItemArray);
                                insertTableToSet(structure, newTable, index);
                            }
                        }
                    }
                }
            }
            else
            {
                structure = new DataSet();
                for (int i = 0; i < newSheets.Tables.Count; i++)
                {
                    DataTable dt = newSheets.Tables[i];
                    // web games
                    if ((gameConfig[dt.TableName])[1] == 0)
                    {
                        for (int j = 2; j < dt.Rows.Count; j++)
                        {
                            string newTableName = dt.TableName.ToString() + "-" + ((dynamic)dt.Rows[j])[0]["text"];
                            if (structure.Tables.Contains(newTableName))
                            {
                                structure.Tables[newTableName].Rows.Add(dt.Rows[j].ItemArray);
                                pageIndexOfThisRow = structure.Tables.IndexOf(newTableName);
                                SlidesEditer.addRow(pptPrest, pageIndexOfThisRow + 2 + gamesCount, dt.Rows[j]);
                            }
                            else
                            {
                                DataTable newTable = new DataTable(); // in or out of for sentance ?
                                for (int k = 0; k < dt.Columns.Count; k++)
                                {
                                    DataColumn column = new DataColumn();
                                    column.DataType = Type.GetType("System.Object");
                                    column.ColumnName = ((dynamic)dt.Rows[0])[k]["text"];
                                    newTable.Columns.Add(column);
                                }
                                newTable.TableName = newTableName;
                                newTable.Rows.Add(dt.Rows[0].ItemArray);
                                newTable.Rows.Add(dt.Rows[j].ItemArray);
                                if (gameList[gameList.Count - 1] == dt.TableName)
                                {
                                    structure.Tables.Add(newTable);
                                }
                                else
                                {
                                    for (int a = 0; a < structure.Tables.Count; a++)
                                    {
                                        if (structure.Tables[a].TableName.StartsWith(dt.TableName))
                                        {
                                            pageIndexOfThisRow = structure.Tables.Count;
                                            for(int b = a; b < structure.Tables.Count; b++)
                                            {
                                                if (!structure.Tables[a].TableName.StartsWith(dt.TableName))
                                                {
                                                    pageIndexOfThisRow = b;
                                                    break;
                                                }
                                            }
                                            break;
                                        }
                                    }
                                    structure = insertTableToSet(structure, newTable, pageIndexOfThisRow);
                                }
                                SlidesEditer.addSilde(pptPrest, pageIndexOfThisRow + 2 + gamesCount, newTable.TableName, dt.Rows[0], dt.Rows[j], gameList.IndexOf(dt.TableName));
                                //SlidesEditer.addRow(pptPrest, j + gamesCount, dt.Rows[j]);
                            }

                        }
                    }
                    else
                    //mobile games
                    {
                        int rowCount = dt.Rows.Count;
                        Dictionary<string, int> channelCountDict = new Dictionary<string, int>();
                        for (int j = 2; j < rowCount; j++)
                        {
                            string newTableName = dt.TableName.ToString() + "-" + ((dynamic)dt.Rows[j])[1]["text"];
                            Console.WriteLine("line {0}, {1}", j + 1, newTableName);
                            if (channelCountDict.Keys.Contains(newTableName))
                            {
                                Console.WriteLine("line {0}'s channel exists.", j + 1);
                                int firstPageIndex = structure.Tables.IndexOf(newTableName);
                                int channelCount = channelCountDict[newTableName];
                                if(channelCount%3 != 0)
                                {
                                    Console.WriteLine("add this line to a page existed");
                                    pageIndexOfThisRow = firstPageIndex + channelCount / 3;
                                    DataTable newTable = structure.Tables[pageIndexOfThisRow];
                                    newTable.Rows.Add(dt.Rows[j].ItemArray);
                                    SlidesEditer.addRow(pptPrest, pageIndexOfThisRow +2 + gamesCount, dt.Rows[j]);
                                }
                                else
                                {         
                                    Console.WriteLine("but it still have to build a new page");
                                    DataTable newTable = new DataTable();
                                    newTable.TableName = newTableName + "[" + (channelCount / 3 + 1).ToString() + "]";
                                    for (int k = 0; k < dt.Columns.Count; k++)
                                    {
                                        DataColumn column = new DataColumn();
                                        column.DataType = Type.GetType("System.Object");
                                        column.ColumnName = ((dynamic)dt.Rows[0])[k]["text"];
                                        newTable.Columns.Add(column);
                                    }
                                    newTable.Rows.Add(dt.Rows[0].ItemArray);
                                    newTable.Rows.Add(dt.Rows[j].ItemArray);
                                    pageIndexOfThisRow = firstPageIndex + channelCount / 3;
                                    structure = insertTableToSet(structure, newTable, pageIndexOfThisRow);
                                    SlidesEditer.addSilde(pptPrest, pageIndexOfThisRow + 2 + gamesCount, newTable.TableName, dt.Rows[0], dt.Rows[j], gameList.IndexOf(dt.TableName));
                                    //SlidesEditer.addRow(pptPrest, j + gamesCount, dt.Rows[j]);
                                }
                                channelCountDict[newTableName] += 1;
                            }
                            else
                            {
                                Console.WriteLine("line {0} does not exist, let's build a new page", j+1);
                                DataTable newTable = new DataTable();
                                newTable.TableName = newTableName;
                                for (int k = 0; k < dt.Columns.Count; k++)
                                {
                                    DataColumn column = new DataColumn();
                                    column.DataType = Type.GetType("System.Object");
                                    column.ColumnName = ((dynamic)dt.Rows[0])[k]["text"];
                                    newTable.Columns.Add(column);
                                }
                                newTable.Rows.Add(dt.Rows[0].ItemArray);
                                newTable.Rows.Add(dt.Rows[j].ItemArray);
                                channelCountDict.Add(newTableName, 1);
                                //structure.Tables.Add(newTable);
                                pageIndexOfThisRow = structure.Tables.Count;
                                if(gameList[gameList.Count-1] == dt.TableName)
                                {
                                    structure.Tables.Add(newTable);
                                }
                                else
                                {
                                    for (int a = 0; a < structure.Tables.Count; a++)
                                    {
                                        if (structure.Tables[a].TableName.StartsWith(dt.TableName))
                                        {
                                            pageIndexOfThisRow = structure.Tables.Count;
                                            for (int b = a; b < structure.Tables.Count; b++)
                                            {
                                                if (!structure.Tables[a].TableName.StartsWith(dt.TableName))
                                                {
                                                    pageIndexOfThisRow = b;
                                                    break;
                                                }
                                            }
                                            break;
                                        }
                                    }
                                    structure = insertTableToSet(structure, newTable, pageIndexOfThisRow);
                                }
                                SlidesEditer.addSilde(pptPrest, pageIndexOfThisRow + 2 + gamesCount, newTable.TableName, dt.Rows[0], dt.Rows[j],  gameList.IndexOf(dt.TableName));
                                //SlidesEditer.addRow(pptPrest, j + gamesCount, dt.Rows[j]);
                                Console.WriteLine("line {0}'s new page has been added to dataset", j + 1);
                            }
                        }                   
                    }
                        
                }
                
            }
            string json = JsonConvert.SerializeObject(structure, Formatting.Indented);
            File.WriteAllText(structureFile, json);
            return structure;
        }

        public static List<object> getSlidesIndex(DataSet structure)
        {
            List<object> slidesIndex = new List<object>();
            List<string> games = new List<string>();
            List<string> material = new List<string>();
            List<int> rowCount = new List<int>();
            for (int i = 0; i < structure.Tables.Count; i++)
            {
                games[i] = structure.Tables[i].TableName.Split(new char[1] { '-' })[0];
                material[i] = structure.Tables[i].TableName.Split(new char[1] { '-' })[1];
                rowCount[i] = structure.Tables[i].Rows.Count - 1;
            }
            slidesIndex.Add(games);
            slidesIndex.Add(material);
            slidesIndex.Add(rowCount);
            return slidesIndex;
        }

        public static DataSet jsonToStructure(string slidesMapJson)
        {
            DataSet structure = new DataSet();
            SortedDictionary<string, object> structureDict = JsonConvert.DeserializeObject<SortedDictionary<string, object>>(slidesMapJson);
            foreach(KeyValuePair<string, object> kvp in structureDict)
            {
                Console.WriteLine("key is {0}, value type is {1}", kvp.Key, kvp.Value.GetType());
                //Console.WriteLine(kvp.Value);
                DataTable a = JsonConvert.DeserializeObject<DataTable>(kvp.Value.ToString());
            }
            return structure;
        }

        static void Main(string[] args)
        {
            
            initialize();
            string excelName = Environment.CurrentDirectory + "\\a.xlsx";
            
            
            string pptName = Environment.CurrentDirectory + "\\test.pptx";
            PowerPoint.Presentation ppt = SlidesEditer.openPPT(pptName);
            DataSet sheets = ReadExcel(excelName, gameConfig);
            
            //string rawJson = File.ReadAllText(@"SlidesMap.json");
            //DataSet structure = jsonToStructure(rawJson);
            makeStructure(ppt,sheets);
            Console.WriteLine("Finish");
            Console.ReadKey();
        }
    }
}
