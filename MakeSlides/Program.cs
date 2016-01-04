using System;
using System.Data;
using System.IO;
using Excel;
using PowerPointOperator;
using Newtonsoft.Json;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Microsoft.Office.Core;

class Program
{
    private static Dictionary<string, int[]> gameConfig = new Dictionary<string, int[]>();
    private static int gamesCount = 0;
    private static List<string> gameList = new List<string>();
    private static string configFile = Environment.CurrentDirectory + "\\Config.json";
    private static string projectFolder = Environment.CurrentDirectory + "\\Project";
    private static string newProjectFolder = String.Empty;
    private static string projectSlides = projectFolder + "\\Sample.pptx";
    private static string excelsHere = projectFolder + "\\ExcelsHere";
    private static string tempoFile = projectFolder + "\\tempo.txt";
    //private static string outputFolder = "OutPut";
    private static string structureFile = String.Empty;
    //private static DataSet structure = new DataSet();

    public static Boolean initialize()
    {
        if (File.Exists(configFile))
        {
            gameConfig = getConfigFile(configFile);
            if(gameConfig == null)
            {
                Console.WriteLine("--> Config.json should be configured correctly.");
                return false;
            }
            gamesCount = gameConfig.Keys.Count;
            foreach (string game in gameConfig.Keys)
            {
                gameList.Add(game);
            }
        }
        else
        {
            File.Create(configFile).Dispose();
            Console.WriteLine("--> A correct config file is needed.");
            Console.ReadKey();
            return false;
        }
        if (!File.Exists("Sample.pptx"))
        {
            Console.WriteLine("--> No Sample.pptx in root directory.");
            Console.ReadKey();
            return false;
        }
        if (!File.Exists(projectFolder))
        {
            Directory.CreateDirectory(projectFolder);
        }
        if (!File.Exists(excelsHere))
        {
            Directory.CreateDirectory(excelsHere);
        }
        if (!File.Exists(projectSlides))
        {
            File.Copy("Sample.pptx", projectSlides);
        }
        if (!File.Exists(tempoFile))
        {
            File.CreateText(tempoFile).Dispose();
        }
        return true;
    }

    public static Dictionary<string, int[]> getConfigFile(string configFilePath)
    {
        Dictionary<string, int[]> customization = new Dictionary<string, int[]>();
        string rawJson = File.ReadAllText(@"Config.json");
        if(rawJson != String.Empty)
        {
            try
            {
                customization = JsonConvert.DeserializeObject<Dictionary<string, int[]>>(rawJson);
            }
            catch (Exception ex)
            {
                customization = null;
                Console.WriteLine("--> Error: {0}.", ex);
            }
        }
        else
        {
            customization = null;
        }
        return customization;
    }

    public static int RGBToIntBGR(string rgb)
    {
        int rr, gg, bb;
        rr = gg = bb = 0;
        if(rgb.Length > 0 & rgb.Length <= 2)
        {
            string r = rgb.Substring(0, rgb.Length);
            rr = Int32.Parse(r, System.Globalization.NumberStyles.HexNumber);
        }
        else if(rgb.Length > 2 & rgb.Length <=4)
        {
            string g = rgb.Substring(2, rgb.Length - 2);
            gg = Int32.Parse(g, System.Globalization.NumberStyles.HexNumber);
        }
        else if(rgb.Length > 4 & rgb.Length <= 6)
        {
            string b = rgb.Substring(4, rgb.Length - 4);
            bb = Int32.Parse(b, System.Globalization.NumberStyles.HexNumber);
        }
        else if(rgb.Length > 6)
        {
            string b = rgb.Substring(4, 2);
            bb = Int32.Parse(b, System.Globalization.NumberStyles.HexNumber);
        }

        return bb * 10 ^ 4 + gg * 10 ^ 2 + rr;
    }

    public static void regulateData(DataTable dt, int width)
    {
		if(dt.Columns.Count > width)
        {
			for(int i = width; i < dt.Columns.Count; i ++)
            {
				dt.Columns.RemoveAt(i);
			}
		}
        else
        {
            width = dt.Columns.Count;
        }

        foreach (DataRow dr in dt.Rows)
        {
            for (int i = 0; i < width; i++)
            {
                if(!string.IsNullOrWhiteSpace(((dynamic)dr[i])["color"]))
                {
                    Console.WriteLine("color is {0}:", ((dynamic)dr[i])["color"]);
                    ((dynamic)dr[i])["color"] = RGBToIntBGR(((dynamic)dr[i])["color"]);
                    Console.WriteLine("color2 is {0}:", ((dynamic)dr[i])["color"]);
                }
                else
                {
                    ((dynamic)dr[i])["color"] = string.Empty;
                }
                string text = ((dynamic)dr[i])["text"];
                dr[i] = new Dictionary<string, object> { { "text", ((dynamic)dr[i])["text"] }, { "color", ((dynamic)dr[i])["color"] }};
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
        DataSet sheets = ExcelReader.getAllSheets(excelFile);
        if (sheets != null)
        {
            if (gameConfig != null)
            {
                for (int i = sheets.Tables.Count - 1; i >= 0; i--)
                {
                    if (sheets.Tables[i].Rows.Count < 3)
                    {
                        sheets.Tables.Remove(sheets.Tables[i]);
                        continue;
                    }
                    string tableName = sheets.Tables[i].TableName;
                    if (!gameConfig.Keys.Contains(tableName))
                    {
                        sheets.Tables.Remove(sheets.Tables[i]);
                    }
                    else
                    {
                        int width = ((dynamic)gameConfig[tableName])[2];
                        regulateData(sheets.Tables[i], width);
                    }
                }
            }
            else
            {
                foreach (DataTable dt in sheets.Tables)
                {
                    regulateData(dt, dt.Columns.Count);
                }
            }
            // export excel files to json
            /*
            json = JsonConvert.SerializeObject(sheets, Formatting.Indented);
            Console.WriteLine("--Data is being written to json file--");
            File.WriteAllText(excelFile + @".json", json);
            Console.WriteLine("--Write operation is complete--");
            */
        }
        return sheets;
    }

    public static DataSet insertTableToSet(DataSet ds, DataTable dt, int index)
    {
        if (index < 0 | index >= ds.Tables.Count)
        {
            ds.Tables.Add(dt);
            return ds;
        }
        else
        {
            DataSet newDataSet = new DataSet();
            for (int i = 0; i < index; i++)
            {
                newDataSet.Tables.Add(ds.Tables[i].Copy());
            }
            newDataSet.Tables.Add(dt);
            for (int i = index; i < ds.Tables.Count; i++)
            {
                newDataSet.Tables.Add(ds.Tables[i].Copy());
            }
            return newDataSet;
        }
    }

    public static DataSet makeStructure(PowerPoint.Presentation pptPrest, DataSet newSheets, DataSet structure)
    {
        Console.WriteLine("--> Processing...");
        List<string> slidesIndex = getSlidesIndex(structure);
        for (int i = 0; i < newSheets.Tables.Count; i++)
        {
            DataTable dt = newSheets.Tables[i];
            for (int j = 2; j < dt.Rows.Count; j++)
            {
                if (((dynamic)dt.Rows[j][0])["text"] == String.Empty)
                {
                    continue;
                }
                slidesIndex = getSlidesIndex(structure);
                string newTableName = dt.TableName.ToString() + "-" + ((dynamic)dt.Rows[j])[0]["text"];
                int firstPageOfGame = slidesIndex.FindIndex(param => param.Equals(dt.TableName));
                int lastPageOfGame = slidesIndex.FindLastIndex(param => param.Equals(dt.TableName));
                if (firstPageOfGame != -1)
                {
                    bool pageFound = false;
                    for (int k = lastPageOfGame; k >= firstPageOfGame; k--)
                    {
                        if (structure.Tables[k].TableName.StartsWith(newTableName))
                        {
                            pageFound = true;
                            string foundPageName = structure.Tables[k].TableName;
                            //web game
                            if ((gameConfig[dt.TableName])[1] == 0)
                            {
                                structure.Tables[k].Rows.Add(dt.Rows[j].ItemArray);
                                SlidesEditer.addRow(pptPrest, k + 2 + gamesCount, dt.Rows[j]);
                                break;
                            }
                            //mobile game
                            else
                            {
                                if (structure.Tables[k].Rows.Count > 3)
                                {
                                    string regex = @"((\d+))";
                                    Match channelIndex = Regex.Match(foundPageName, regex);
                                    if (channelIndex.Length == 0)
                                    {
                                        newTableName += "(2)";
                                    }
                                    else
                                    {
                                        newTableName += "(" + (Convert.ToInt32(channelIndex.Groups[1].Value) + 1).ToString() + ")";
                                    }
                                    DataTable newTable = dt.Clone();
                                    newTable.TableName = newTableName;
                                    newTable.Rows.Add(dt.Rows[0].ItemArray);
                                    newTable.Rows.Add(dt.Rows[j].ItemArray);
                                    structure = insertTableToSet(structure, newTable, k + 1);
                                    SlidesEditer.addSilde(pptPrest, k + 3 + gamesCount, newTable.TableName, dt.Rows[0], dt.Rows[j], gameList.IndexOf(dt.TableName));
                                    break;
                                }
                                else
                                {
                                    structure.Tables[k].Rows.Add(dt.Rows[j].ItemArray);
                                    SlidesEditer.addRow(pptPrest, k + 2 + gamesCount, dt.Rows[j]);
                                    break;
                                }
                            }
                        }
                    }
                    if (pageFound == false)
                    {
                        DataTable newTable = dt.Clone();
                        newTable.TableName = newTableName;
                        newTable.Rows.Add(dt.Rows[0].ItemArray);
                        newTable.Rows.Add(dt.Rows[j].ItemArray);
                        structure = insertTableToSet(structure, newTable, lastPageOfGame + 1);
                        SlidesEditer.addSilde(pptPrest, lastPageOfGame + 3 + gamesCount, newTable.TableName, dt.Rows[0], dt.Rows[j], gameList.IndexOf(dt.TableName));
                        continue;
                    }
                }
                else
                // current game does not exist;
                {
                    if (structure.Tables.Count == 0)
                    {
                        DataTable newTable = dt.Clone();
                        newTable.TableName = newTableName;
                        newTable.Rows.Add(dt.Rows[0].ItemArray);
                        newTable.Rows.Add(dt.Rows[j].ItemArray);
                        structure.Tables.Add(newTable);
                        SlidesEditer.addSilde(pptPrest, 2 + gamesCount, newTable.TableName, dt.Rows[0], dt.Rows[j], gameList.IndexOf(dt.TableName));
                    }
                    else
                    {
                        int insertIndex = 0;
                        for (int a = structure.Tables.Count - 1; a >= 0; a--)
                        {
                            string pageGame = ((structure.Tables[a].TableName).Split(new char[1] { '-' }))[0];
                            int pageGameIndex = ((dynamic)gameConfig[pageGame])[0];
                            int indexOfGame = ((dynamic)gameConfig[dt.TableName])[0];
                            if (indexOfGame > pageGameIndex)
                            {
                                insertIndex = a + 1;
                                break;
                            }
                        }
                        DataTable newTable = dt.Clone();
                        newTable.TableName = newTableName;
                        newTable.Rows.Add(dt.Rows[0].ItemArray);
                        newTable.Rows.Add(dt.Rows[j].ItemArray);
                        structure = insertTableToSet(structure, newTable, insertIndex);
                        SlidesEditer.addSilde(pptPrest, insertIndex + 2 + gamesCount, newTable.TableName, dt.Rows[0], dt.Rows[j], gameList.IndexOf(dt.TableName));
                        continue;
                    }
                }
            }
        }
        //export slidemaps to json
        /*
        string json = JsonConvert.SerializeObject(structure, Formatting.Indented);
        File.WriteAllText(structureFile, json);
        */
        pptPrest.SaveAs(projectSlides);
        Console.WriteLine("--> Mission success."); 
        Console.WriteLine("--> Backup data...");
        if (File.Exists(projectFolder + "\\" + "SlidesMap.xlsx"))
        {
            File.Delete(projectFolder + "\\" + "SlidesMap.xlsx");
        }
        ExcelWriter.ExportDataSet(structure, projectFolder + "\\" + "SlidesMap.xlsx");
        Console.WriteLine("--> Backup Completed.");
        Console.WriteLine("--");
        return structure;
    }

    public static List<string> getSlidesIndex(DataSet structure)
    {
        List<string> games = new List<string>();
        for (int i = 0; i < structure.Tables.Count; i++)
        {
            games.Add((structure.Tables[i].TableName.Split(new char[1] { '-' }))[0]);
        }
        return games;
    }

    private static void getLists(out IEnumerable<string> excelFiles, out IEnumerable<string> tempo)
    {
        excelFiles = Directory.EnumerateFiles(excelsHere, "*.*", SearchOption.AllDirectories)
            .Where(s => s.EndsWith(".xlsx"));
        tempo = File.ReadLines(tempoFile);
    }

    static void Main(string[] args)
    {      
        Boolean bl = initialize();

        while (!bl)
        {
            Console.WriteLine("--> Press Any Key to Fresh.");
            ConsoleKeyInfo input = Console.ReadKey();
            if (!string.IsNullOrEmpty(input.KeyChar.ToString()))
            {
                Console.Clear();
                bl = initialize();
            }  
        }

        DataSet structure = new DataSet();

        IEnumerable<string> excelFiles;
        IEnumerable<string> tempo;

        getLists(out excelFiles, out tempo);

        List<string> todoList = new List<string>();
        foreach (string s in excelFiles)
        {
            if (!tempo.Contains<string>(Path.GetFileName(s)))
            {
                todoList.Add(s);
            }
        }

        while(todoList.Count<string>() == 0)
        {
            Console.WriteLine("--> No new excel files, Press any key to fresh.");
            Console.ReadKey();
            Console.Clear();
            initialize();
            getLists(out excelFiles, out tempo);
            foreach (string s in excelFiles)
            {
                if (!tempo.Contains<string>(Path.GetFileName(s)))
                {
                    todoList.Add(s);
                }
            }
        }
        PowerPoint.Application app = new PowerPoint.Application();
        PowerPoint.Presentation ppt = SlidesEditer.openPPT(projectSlides, app);
        /*
        try
        {
            ppt.Windows._Index(0);
        }
        catch
        {
            ppt.Close();
            app.Quit();
            GC.Collect();
        }
        */
        if (todoList.Count<string>() != 0)
        {
            Console.WriteLine("-->  _(:3 」∠)_ ..");
            foreach (string s in todoList)
            {
                Console.WriteLine("--> Reading {0}:", Path.GetFileName(s));
                DataSet sheets = ReadExcel(s, gameConfig);
                if (tempo.Count<string>() != 0)
                {
                    structureFile = projectFolder + "\\SlidesMap.xlsx";
                    structure = ExcelReader.getAllSheets(structureFile);
                    for (int i = 0; i < structure.Tables.Count; i++)
                    {
                        regulateData(structure.Tables[i], structure.Tables[i].Columns.Count);
                    }
                }       
                makeStructure(ppt, sheets, structure);
                using (StreamWriter sw = File.AppendText(tempoFile))
                {
                    sw.WriteLine(Path.GetFileName(s));
                    sw.Close();
                }
            }
            
        }
        Console.WriteLine("--> Done.");
        Console.WriteLine("--> Press Enter to Exit.");
        ppt.Close();
        app.Quit();
        GC.Collect();
        Console.SetWindowPosition(0, 0);
        Console.ReadKey();
    }
}
