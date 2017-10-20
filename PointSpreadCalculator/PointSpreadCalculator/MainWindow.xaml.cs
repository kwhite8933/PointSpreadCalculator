using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;

namespace PointSpreadCalculator
{

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        string currentWeek = ""; // Current week being processed
        int gamesThisWeek = 0;
        Dictionary<string, string> Teams = new Dictionary<string, string>()
        {{"49ers", ""},
         {"Bears", ""},
         {"Bengals", ""},
         {"Bills", ""},
         {"Broncos", ""},
         {"Browns", ""},
         {"Buccaneers", ""},
         {"Cardinals", ""},
         {"Chargers", ""},
         {"Chiefs", ""},
         {"Colts", ""},
         {"Cowboys", ""},
         {"Dolphins", ""},
         {"Eagles", ""},
         {"Falcons", ""},
         {"Giants", ""},
         {"Jaguars", ""},
         {"Jets", ""},
         {"Lions", ""},
         {"Packers", ""},
         {"Panthers", ""},
         {"Patriots", ""},
         {"Raiders", ""},
         {"Rams", ""},
         {"Ravens", ""},
         {"Redskins", ""},
         {"Saints", ""},
         {"Seahawks", ""},
         {"Steelers", ""},
         {"Texans", ""},
         {"Titans", ""},
         {"Vikings", ""}};
        
        
        static readonly string[] cities = new string[32] { "Arizona", "Atlanta", "Baltimore", "Buffalo", "Carolina", "Chicago", "Cincinnati", "Cleveland",
                                           "Dallas", "Denver", "Detroit", "Green Bay", "Houston", "Indianapolis", "Jacksonville", 
                                           "Kansas City", "Los Angeles (LAC)", "Los Angeles (LAR)", "Miami", "Minnesota", "New England",
                                           "New Orleans", "New York (NYG)", "New York (NYJ)", "Oakland", "Philadelphia", "Pittsburgh",
                                           "San Francisco", "Seattle", "Tampa Bay", "Tennessee", "Washington"};

        public MainWindow()
        {
            InitializeComponent();
        }

        #region Click Handlers
        /////////////////////////////////////////////////////// Click Handlers ///////////////////////////////////////////////////////

        /// <summary>
        /// Calculates the spread for each team and displays those spreads in a word or excel documenet
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCalculate_Click(object sender, RoutedEventArgs e)
        {
            //createWordTable();
            createExcelTable();
        }

        /// <summary>
        /// Converts each city to it's team name and prints it on the command line in alphabetical order
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnConvert_Click(object sender, RoutedEventArgs e)
        {
            string[] temp = new string[cities.Length];
            for (int i = 0; i < cities.Length; i++)
            {
                temp[i] = convertCityToTeam(cities[i]);
            }
            Array.Sort(temp);
            foreach (string name in temp)
            {
                Console.WriteLine(name);
            }
        }

        /// <summary>
        /// Reads in a file of each game and the point spread for each respective team.
        /// The games will be parced and properly processed into it's individual parts.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnReadFile_Click(object sender, RoutedEventArgs e)
        {
            currentWeek = tbCurrentWeek.Text.ToString();
            Console.WriteLine(currentWeek);
            readFile(string.Format("C:\\Users\\Kyle\\Documents\\Code\\C#\\PointSpreadCalculator\\PointSpreadCalculator\\Matchups\\Week{0}.txt", currentWeek));
            //readFile(string.Format("Z:\\PointSpreadCalculator\\Matchups\\Week{0}.txt", currentWeek));
        }
        #endregion

        /////////////////////////////////////////////////////// Member Functions ///////////////////////////////////////////////////////

        //private void createWordTable()
        //{
        //    // initialize the number of rows and columns in the table
        //    int rows = gamesThisWeek*2 + 1;
        //    int columns = 2;
        //    object objMissing = System.Reflection.Missing.Value;
        //    object objEOF = "\\endofdoc";

        //    // initialize microsoft word document object
        //    Microsoft.Office.Interop.Word.Application appObj;
        //    Microsoft.Office.Interop.Word.Document docObj;
        //    appObj = new Microsoft.Office.Interop.Word.Application();
        //    appObj.Visible = true;
        //    docObj = appObj.Documents.Add(ref objMissing, ref objMissing, ref objMissing, ref objMissing);
        //    Microsoft.Office.Interop.Word.Table tableObj;
        //    Microsoft.Office.Interop.Word.Range rangeOfWord = docObj.Bookmarks.get_Item(objEOF).Range;
        //    tableObj = docObj.Tables.Add(rangeOfWord, rows, columns, ref objMissing, ref objMissing);

        //    // set default formatting for table
        //    tableObj.Range.ParagraphFormat.SpaceAfter = 8;
        //    tableObj.Range.Font.Name = "Times New Roman";
        //    tableObj.Range.Font.Size = 9;
        //    tableObj.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

        //    //tableObj.PreferredWidth = 100;
        //    // Display table headers
        //    tableObj.Cell(1, 1).Range.Text = "Name";
        //    tableObj.Cell(1, 2).Range.Text = "Spread";

        //    // Populate table with values
        //    string str;

        //    for (int i = 2; i <= rows; i++)
        //    {
        //        for (int j = 1; j <= columns; j++)
        //        {
        //            try
        //            {
        //                if (j == 1)
        //                {
        //                    str = convertCityToTeam(lstTeams[i-2].getName());
        //                }
        //                else
        //                {
        //                    str = lstTeams[i-2].getSpread().ToString();
        //                }
        //                //int currRow = i - 1;
        //                //str = lstTeams[currRow].getName();

        //                tableObj.Cell(i, j).Range.Text = str;

        //            }
        //            catch (System.Runtime.InteropServices.COMException e)
        //            {
        //                Console.WriteLine("Caught Exception: " + e);
        //                break;
        //            }
        //        }
        //    }
        //    tableObj.Rows[1].Range.Font.Bold = 1;
        //    tableObj.Rows[1].Range.Font.AllCaps = 1;

        //    tableObj.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
        //    tableObj.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
        //    tableObj.Range.Rows[1].Shading.BackgroundPatternColor = WdColor.wdColorGray25;
        //    tableObj.Range.Rows[1].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleDouble;
        //    tableObj.Range.Rows.HeightRule = WdRowHeightRule.wdRowHeightAuto;
        //}

        /// <summary>
        /// Generates an excel table displaying the spread for each team
        /// </summary>
        private void createExcelTable()
        {
            // initialize the amount of rows and columns
            int rows = gamesThisWeek * 2 + 1;
            int columns = 2;

            // initialize the excel workbook
            Microsoft.Office.Interop.Excel.Application _xl;
            Microsoft.Office.Interop.Excel._Workbook _wb;
            Microsoft.Office.Interop.Excel._Worksheet _sheet;
            //Microsoft.Office.Interop.Excel.Range _range;
            object misValue = System.Reflection.Missing.Value;
            try
            {
                // start excel and get Application object
                _xl = new Microsoft.Office.Interop.Excel.Application();
                _xl.Visible = true;

                // get a new workbook
                _wb = (Microsoft.Office.Interop.Excel._Workbook)(_xl.Workbooks.Add(""));
                _sheet = (Microsoft.Office.Interop.Excel._Worksheet)_wb.ActiveSheet;

                // Add table headers
                _sheet.Cells[1, 1] = "Name";
                _sheet.Cells[1, 2] = "Spread";

                //lstTeams.Sort((x, y) => string.Compare(convertCityToTeam(x.getName()), convertCityToTeam(y.getName())));

                // Format cells
                _sheet.get_Range("A1", "D1").Font.Bold = true;
                _sheet.get_Range("A1", "D1").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                _sheet.get_Range("B2", "B33").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                // Populate Cells
                string str;
                int row = 2;
                foreach (KeyValuePair<string, string> team in Teams)
                {
                    for (int col = 1; col <= columns; col++)
                    {
                        try
                        {
                            str = (col == 1) ? team.Key : team.Value;
                            _sheet.Cells[row, col] = str.Trim();
                        }
                        catch(System.Runtime.InteropServices.COMException e)
                        {
                            Console.WriteLine("Caught Exception: {0}", e);
                        }
                    }
                    row++;
                }
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                Console.WriteLine("Caught Exception: " + e);
            }
            Console.WriteLine("Successfully generated excel table");
            this.Close();
        }

        /// <summary>
        /// Reads in a file of each game and the point spread for each respective team.
        /// The games will be parced and properly processed into it's individual parts.
        /// </summary>
        /// <param name="_file">The file containing the list of games for that week</param>
        private void readFile(string _file)
        {
            // Make sure the file exists
            if (System.IO.File.Exists(_file))
            {
                string[] lines = { };
                string[] splitGame = { };

                // Read in all lines to array
                // The number of lines is the number of games in that week
                lines = System.IO.File.ReadAllLines(_file);
                gamesThisWeek = lines.Length;
                for (int i = 0; i < lines.Length; i++)
                {
                    lines[i] = parceGame(lines[i]);
                    Console.WriteLine(lines[i]);
                }

                // Parce each game to get the team name and that teams' spread for the week
                for (int i = 0; i < lines.Length; i++)
                {
                    // .Split(null) splits each line of the file by the whitespace
                    splitGame = lines[i].Split('-'); // splits each line of the file by the '-'
                    splitGame[2] = splitGame[2].Trim();
                    // if the team name contains an '@', that team is the home team
                    // otherwise that team is the away team
                    if (splitGame[0].Contains("@"))
                    {
                        string key0 = convertCityToTeam(splitGame[0].Substring(1, splitGame[0].Length - 1));
                        string key1 = convertCityToTeam(splitGame[1]);
                        Teams[key0] = (splitGame[2].Equals("0.0")) ? splitGame[2] : "-" + splitGame[2];
                        Teams[key1] = splitGame[2];
                    }
                    else if (splitGame[1].Contains("@"))
                    {
                        string key0 = convertCityToTeam(splitGame[0]);
                        string key1 = convertCityToTeam(splitGame[1].Substring(1, splitGame[1].Length - 1));
                        Teams[key0] = (splitGame[2].Equals("0.0")) ? splitGame[2] : "-" + splitGame[2];
                        Teams[key1] = splitGame[2];
                    }
                    else
                    {
                        Console.WriteLine("Error finding home team, check games file to make sure a home team is declared");
                    }

                }
                //Array.Sort(lines);
                //System.IO.File.WriteAllLines("C:\\Users\\Kyle\\Documents\\Code\\C#\\PointSpreadCalculator\\output.txt", lines);
                //Console.WriteLine(convertCityToTeam(lstGames[8].getHomeTeam()));
                //System.IO.File.WriteAllLines("Z:\\PointSpreadCalculator\\output.txt", lines);
                Console.WriteLine("Successfully Read File");

            }
            else
            {
                Console.WriteLine("File does not exist");
                this.Close();
            }
        }

        /// <summary>
        /// Parces the raw data file of each game and formats it so data can be extracted from it.
        /// </summary>
        /// <param name="_game">The unformatted line containing all of the game's information</param>
        /// <returns></returns>
        private string parceGame(string _game)
        {
            string result = "";
            if(_game.Contains("--")){
                _game = _game.Replace("--", "0.0");
            }

            Match match = getMatch("(thu|sun|mon).*", _game); // Removes the date the game is happening on
            result = _game.Substring(0, match.Index);
            
            match = getMatch(@"(\d\d?.\d\s){3}", result); // Removes the spreads from the previous days during the week
            result = result.Substring(0, match.Index) + result.Substring(match.Index + match.Value.Length);

            match = getMatch(@"\s+vs\s+", result); // Replaces the *team* vs *team* with a '-' for easier parcing later
            result = result.Substring(0, match.Index) + "-" + result.Substring(match.Index + match.Value.Length);

            match = getMatch(@"\t", result); // Replaces the space after the second team with a '-' for eaiser parcing later
            result = result.Substring(0, match.Index) + "-" + result.Substring(match.Index + match.Value.Length);
            //.WriteLine("{0}", result);

            return result;
        }

        private Match getMatch(string _pattern, string input)
        {
            //string pattern = _pattern;
            Regex reg = new Regex(_pattern, RegexOptions.IgnoreCase);
            Match match = reg.Match(input);
            return match;
        }

        /// <summary>
        /// Converts a city to that city's team name
        /// </summary>
        /// <param name="city">The city that is to be converted to a team name</param>
        /// <returns></returns>
        private string convertCityToTeam(string city)
        {
            string name = "";
            switch (city)
            {
                case "Arizona":
                    name = "Cardinals";
                    break;
                case "Atlanta":
                    name = "Falcons";
                    break;
                case "Baltimore":
                    name = "Ravens";
                    break;
                case "Buffalo":
                    name = "Bills";
                    break;
                case "Carolina":
                    name = "Panthers";
                    break;
                case "Chicago":
                    name = "Bears";
                    break;
                case "Cincinnati":
                    name = "Bengals";
                    break;
                case "Cleveland":
                    name = "Browns";
                    break;
                case "Dallas":
                    name = "Cowboys";
                    break;
                case "Denver":
                    name = "Broncos";
                    break;
                case "Detroit":
                    name = "Lions";
                    break;
                case "Green Bay":
                    name = "Packers";
                    break;
                case "Houston":
                    name = "Texans";
                    break;
                case "Indianapolis":
                    name = "Colts";
                    break;
                case "Jacksonville":
                    name = "Jaguars";
                    break;
                case "Kansas City":
                    name = "Chiefs";
                    break;
                case "Los Angeles (LAC)":
                    name = "Chargers";
                    break;
                case "Los Angeles (LAR)":
                    name = "Rams";
                    break;
                case "Miami":
                    name = "Dolphins";
                    break;
                case "Minnesota":
                    name = "Vikings";
                    break;
                case "New England":
                    name = "Patriots";
                    break;
                case "New Orleans":
                    name = "Saints";
                    break;
                case "New York (NYG)":
                    name = "Giants";
                    break;
                case "New York (NYJ)":
                    name = "Jets";
                    break;
                case "Oakland":
                    name = "Raiders";
                    break;
                case "Philadelphia":
                    name = "Eagles";
                    break;
                case "Pittsburgh":
                    name = "Steelers";
                    break;
                case "San Francisco":
                    name = "49ers";
                    break;
                case "Seattle":
                    name = "Seahawks";
                    break;
                case "Tampa Bay":
                    name = "Buccaneers";
                    break;
                case "Tennessee":
                    name = "Titans";
                    break;
                case "Washington":
                    name = "Redskins";
                    break;
                default:
                    Console.WriteLine("Could not find team name");
                    break;
            }

            return name;
        }
    }
}
