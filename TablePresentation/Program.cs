using System;
using System.Data;
using System.IO;
using ExcelManipulater;
using PowerPointOperator;
using Newtonsoft.Json;
using System.Linq;
using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Text.RegularExpressions;

class Program
{
    private static Dictionary<string, int[]> gameConfig = new Dictionary<string, int[]>();
    private static int gamesCount = 0;
    private static List<string> gameList = new List<string>();
    private static string configFile = Environment.CurrentDirectory + "\\Config.json";
    private static string projectFolder = Environment.CurrentDirectory + "\\Project";
    private static string newProjectFolder = String.Empty;
    private static string projectSlides = String.Empty;
    //private static string backupFolder = "Backup";
    //private static string outputFolder = "OutPut";
    private static string structureFile = String.Empty;
    private static DataSet structure = new DataSet();

    public static void initialize()
    {
        if (File.Exists(configFile))
        {
            gameConfig = getConfigFile(configFile);
            gamesCount = gameConfig.Keys.Count;
            foreach (string game in gameConfig.Keys)
            {
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
        newProjectFolder = String.Empty;
        projectSlides = String.Empty;
        structureFile = String.Empty;
    }

    public static void reset()
    {
        newProjectFolder = String.Empty;
        projectSlides = String.Empty;
        structureFile = String.Empty;
    }

    public static Dictionary<string, int[]> getConfigFile(string configFilePath)
    {
        Dictionary<string, int[]> customization = new Dictionary<string, int[]>();
        string rawJson = File.ReadAllText(@"Config.json");
        //Console.WriteLine(rawJson);
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

                    if (format.IndexOf("%") != -1)
                    {
                        ((dynamic)dr[i])["text"] = (Math.Round(double.Parse(text), 4, MidpointRounding.AwayFromZero) * 100).ToString() + "%";

                    }
                    else
                    {
                        ((dynamic)dr[i])["text"] = Math.Round(double.Parse(text), 2, MidpointRounding.AwayFromZero).ToString();
                    }
                }
                dr[i] = new Dictionary<string, object> { { "text", ((dynamic)dr[i])["text"] }, { "color", ((dynamic)dr[i])["color"] } };
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
        //string json = String.Empty;
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
                    regulateData(dt);
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
                //Console.WriteLine("row {0}, firstpageofgame is {1}, last is {2}", j - 1, firstPageOfGame, lastPageOfGame);
                if (firstPageOfGame != -1)
                {
                    bool pageFound = false;
                    //Console.WriteLine("row{0}", j - 1);
                    for (int k = lastPageOfGame; k >= firstPageOfGame; k--)
                    {
                        //Console.WriteLine("this row is {0}, current page is {1}", newTableName, structure.Tables[k].TableName);
                        if (structure.Tables[k].TableName.StartsWith(newTableName))
                        {
                            pageFound = true;
                            //Console.WriteLine("row{0} found page {1}", j - 1, k + 1);
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
                        Console.WriteLine("game no find, and tables count is 0, {0}", dt.TableName);
                        DataTable newTable = dt.Clone();
                        newTable.TableName = newTableName;
                        newTable.Rows.Add(dt.Rows[0].ItemArray);
                        newTable.Rows.Add(dt.Rows[j].ItemArray);
                        structure.Tables.Add(newTable);
                        SlidesEditer.addSilde(pptPrest, 2 + gamesCount, newTable.TableName, dt.Rows[0], dt.Rows[j], gameList.IndexOf(dt.TableName));
                    }
                    else
                    {
                        Console.WriteLine("game no find, but tables count not 0, {0}", dt.TableName);
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
        string slideMaps = newProjectFolder + "\\" + "SlidesMap-" + DateTime.Now.Hour.ToString() + "-" + DateTime.Now.Minute.ToString() + ".xls";
        if (File.Exists(newProjectFolder + "\\" + "SlidesMap.xls"))
        {
            File.Move(newProjectFolder + "\\" + "SlidesMap.xls", slideMaps);
        }
        ExcelWriter.ExportDataToExcel(structure, newProjectFolder + "\\" + "SlidesMap.xls");
        pptPrest.SaveAs(projectSlides);
        return structure;
    }

    public static List<string> getSlidesIndex(DataSet structure)
    {
        //List<object> slidesIndex = new List<object>();
        List<string> games = new List<string>();
        //List<string> tableNames = new List<string>();
        for (int i = 0; i < structure.Tables.Count; i++)
        {
            games.Add((structure.Tables[i].TableName.Split(new char[1] { '-' }))[0]);
        }
        return games;
    }


    public static void creatNewProject()
    {
        newProjectFolder = Environment.CurrentDirectory + "\\Project" + "\\" + DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + "-" + DateTime.Now.Hour + "-" + DateTime.Now.Minute;
        if (!File.Exists(newProjectFolder))
        {
            Directory.CreateDirectory(newProjectFolder);
        }
        projectSlides = newProjectFolder + "\\" + DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + ".pptx";
        if (!File.Exists(projectSlides))
        {
            File.Copy("Sample.pptx", projectSlides);
        }
        
    }

    public static  void showMenu()
    {
        reset();
        Console.Clear();
        Console.WriteLine("=MainMenu=");
        Console.WriteLine("===============================");
        Console.WriteLine("1.Creat a new project.");
        Console.WriteLine("2.Proceed with last project.");
        Console.WriteLine("3.Enter x to Exit.");
        Console.WriteLine("===============================");
        Console.WriteLine("Enter: ");
        ConsoleKeyInfo input = Console.ReadKey();
        switch (input.KeyChar.ToString())
        {
            case "1":
                //Console.ReadKey();
                creatNewProject();
                Console.Clear();
                Console.WriteLine("===============================");
                Console.WriteLine("Your project has been created.");
                showMenu_1();
                break;
            case "2":
                Console.ReadKey();
                string[] newProjectFolders = Directory.GetDirectories(projectFolder);
                if(newProjectFolders.Length == 0)
                {
                    showMenu();
                }
                else
                {
                    newProjectFolder = newProjectFolders[newProjectFolders.Length - 1];
                    projectSlides = newProjectFolder + "\\" + DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + ".pptx";
                    showMenu_2(1);
                }
                break;
            case "x":
                Console.ReadKey();
                Environment.Exit(0);
                break;
            default:
                showMenu();
                break;
        }
    }

    public static void showMenu_1()
    {
        Console.WriteLine("===============================");
        Console.WriteLine("1.Import a excel file.");
        Console.WriteLine("2.Back to the previous screen.");
        Console.WriteLine("3.Enter x to Exit.");
        Console.WriteLine("===============================");
        ConsoleKeyInfo input = Console.ReadKey();
        switch (input.KeyChar.ToString())
        {
            case "1":
                showMenu_2_1(0);
                break;
            case "2":
                showMenu();
                break;
            case "x":
                Environment.Exit(0);
                break;
            default:
                showMenu_1();
                break;
        }
    }

    public static void showMenu_2(int hasStructure)
    {
        Console.Clear();
        Console.WriteLine("===============================");
        Console.WriteLine("1.Import a excel file.");
        Console.WriteLine("2.Back to the previous screen.");
        Console.WriteLine("3.Enter x to Exit.");
        Console.WriteLine("===============================");
        ConsoleKeyInfo input = Console.ReadKey();
        switch (input.KeyChar.ToString())
        {
            case "1":
                Console.ReadKey();
                showMenu_2_1(hasStructure);
                break;
            case "2":
                Console.ReadKey();
                showMenu();
                break;
            case "x":
                Console.ReadKey();
                Environment.Exit(0);
                break;
            default:
                Console.ReadKey();
                showMenu_2(0);
                break;
        }
    }

    public static void showMenu_2_1(int hasStructure)
    {
        Console.Clear();
        Console.WriteLine("==================================");
        Console.WriteLine("=====drag-and-drop the excel file here.=====");
        Console.WriteLine("==================================");
        string excelName = Console.ReadLine();
        if (!File.Exists(excelName))
        {
            showMenu_2_1(hasStructure);
        }
        DataSet sheets = ReadExcel(excelName, gameConfig);
        PowerPoint.Presentation ppt = SlidesEditer.openPPT(projectSlides);
        if (hasStructure == 1)
        {
            structureFile = newProjectFolder + "\\SlidesMap.xls";
            structure = ExcelReader.ImportDataFromAllSheets(structureFile);
            //Console.WriteLine("has structure, length {0}", structure.Tables.Count);
        }
        makeStructure(ppt, sheets, structure);
        Console.WriteLine("Finish");
        Console.ReadKey();
        showMenu();
    }

    static void Main(string[] args)
    {      
        initialize();
        showMenu();
        Console.ReadKey();
    }
}
