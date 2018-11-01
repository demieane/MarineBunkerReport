using System;
using System.IO;
using System.Data;
using System.Linq;
using LinqToExcel;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace WindowsFormsApp1
{
    public partial class BunkerReport : Form
    {
        //================================================================
        // Initialize: Public variables that all other classes can access
        //================================================================

        // Type of Output
        public bool Weekly = false;
        public bool NonWeekly = false;
        public string AssetName { get; private set; }

        // Excel conditional formatting 
        public float AcceptableRangeFO = new float();
        public float AcceptableRangeDO = new float();
        public bool UserDefinedRangeFO = false;
        public bool UserDefinedRangeDO = false;

        public bool AcceptableBunkerDirectoryPath { get; private set; }
        public bool AcceptableDatabaseDirectoryPath { get; private set; }
        public bool AcceptableChoiceNewUpdate { get; private set; }

        // Bunker Report File Path
        public string BunkerReportDirectoryPath { get; private set; }

        // Database Paths and FileName
        public string CreateNewDatabase { get; private set; }
        public string BunkerDatabaseDirectoryPath { get; private set; }
        public string BunkerDatabaseFilePathWeekly { get; private set; }
        public string BunkerDatabaseFilePathMonthly { get; private set; }

        public string ResultsFilePathWeekly { get; private set; }
        public string ResultsFilePathMonthly { get; private set; }
        public string PathString { get; private set; }

        public string SheetName { get; private set; }
        public string DatabaseFileName { get; private set; }
        public bool DatabaseFileNameState { get; private set; }

        // Data Storage
        public List<string> stringList = new List<string>(); // Βunker report file paths
        public VesselInfo Vessel = new VesselInfo();
        public DatabaseInfo Data = new DatabaseInfo();
        public int Counter = new int();


        // Golden Union Bulk Carrier Fleet 
        public bool eureka = new bool();
        public string[] CurrentFleet;
        public bool[] CurrentFleetStatus; //= new bool[1000];
        // Updating Excel
        public string[] CoordinatesOfLastElement { get; private set; }

        // BUTTONS
        // SELECT SOURCE FOLDER FOR BUNKER REPORTS
        public bool ActionOne = new bool();
        // SELECT DEPOSIT FOLDER FOR NEW DATABASE
        public bool ActionTwo = new bool();
        // SELECT DEPOSIT FOLDER FOR NEW DATABASE
        public bool ActionThree = new bool();

        //================================================================
        // Main Class
        //================================================================

        public BunkerReport()
        {
            // Set conditional Formatting values
            AcceptableRangeFO = 15; // [metric tonnes]
            AcceptableRangeDO = 5;  // [metric tonnes]

            AcceptableBunkerDirectoryPath = false;
            AcceptableDatabaseDirectoryPath = false;
            AcceptableChoiceNewUpdate = false;

            InitializeComponent();

            // Set Labels in Application Window
            label2.Font = new Font(label2.Font, FontStyle.Bold);
            label2.Text = "Fill the following box with the directory containing the Bunker Reports";

            label5.Font = new Font(label2.Font, FontStyle.Bold);
            label5.Text = "Fill the following box with the directory of the .xlsx Database";

            label3.Text = "Directory:";
            label1.Text = "Directory:";
            label4.Text = "Filename:";

            label6.Text = "Acceptable FO Difference [mt]:";
            textBox3.Text = AcceptableRangeFO.ToString();
            label7.Text = "Acceptable LO Difference [mt]:";
            textBox5.Text = AcceptableRangeDO.ToString();

            
        }

        //================================================================
        // Utilities
        //================================================================

        //================================================================
        // Returns: Bunker Report Directory Path
        //================================================================
        public void ChooseFolderReports()
        {
            //ProcessParameters BunkerReport = new ProcessParameters();
            using (var fbd = new FolderBrowserDialog())
            {
                fbd.RootFolder = Environment.SpecialFolder.Desktop;
                fbd.Description = "Select Folder: Weekly Bunker Reports";
                fbd.ShowNewFolderButton = true;
                DialogResult result = fbd.ShowDialog();
                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    // Initial choice for path
                    string path = fbd.SelectedPath;
                    BunkerReportDirectoryPath = path;
                    // Display for user
                    textBox1.Text = fbd.SelectedPath;
                }
            }
        }

        //================================================================
        //  Returns: Bunker Database Directory Path
        //================================================================
        public void ChooseFolderDatabase()
        {
            //ProcessParameters BunkerReport = new ProcessParameters();
            using (var fbd = new FolderBrowserDialog())
            {
                fbd.RootFolder = Environment.SpecialFolder.Desktop;
                fbd.Description = "Select Path: Database";
                fbd.ShowNewFolderButton = true;
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    // Initial choice for path
                    string path = fbd.SelectedPath;
                    BunkerDatabaseDirectoryPath = path;
                    // Display for user
                    textBox2.Text = fbd.SelectedPath;
                }
            }
        }

        //================================================================
        // Returns: Bunker Report File Paths from the given Directory
        //================================================================
        public void FindBunkerReportFilePaths()
        {
            if (File.Exists(BunkerReportDirectoryPath))
            {
                // This path is a file
                MessageBox.Show("This path is a file. Please choose a directory", "Comment", MessageBoxButtons.RetryCancel);
                DialogResult result = DialogResult.Retry;
                
                if (result == DialogResult.Retry)
                {
                    // Regarding bunker report directory
                    if (ActionOne == true || (!string.IsNullOrWhiteSpace(textBox1.Text)))
                    {
                        AcceptableBunkerDirectoryPath = false;
                    }
                    
                }
            }
            else if (Directory.Exists(BunkerReportDirectoryPath))
            {
                int index;
                int sizeOfList;
                index = 0;

                // This path is a directory
                stringList = MultipleFiles.ProcessDirectory(BunkerReportDirectoryPath, index);
                sizeOfList = stringList.Count;
                AcceptableBunkerDirectoryPath = true;
                MultipleFiles.ProcessFile(BunkerReportDirectoryPath);
            }
            else
            {
                MessageBox.Show("Not a valid file or directory", "Comment", MessageBoxButtons.RetryCancel);
                DialogResult result = DialogResult.Retry;
              
                if (result == DialogResult.Retry)
                {
                    AcceptableBunkerDirectoryPath = false;
                }
            }
        }

        //================================================================
        // Creates the final Database and Output Files based on User Info
        //================================================================
        public void DataBaseHandling()
        {
            BunkerReportHandling NewHandlingTools = new BunkerReportHandling();

            // WRITE DATABASE IN .CSV
            if (ActionTwo == true)
            {
                PathString = BunkerDatabaseDirectoryPath + @"\Database\";

                // If the fileName is defined by the user in textBox4
                if (DatabaseFileNameState)
                {
                    if (Weekly)
                    {
                        string pathWrite = PathString + DatabaseFileName + @".xlsx";

                        BunkerDatabaseFilePathWeekly = pathWrite;
                        ResultsFilePathWeekly = PathString + DatabaseFileName + @".txt";
                        Console.WriteLine("BunkerDatabaseFilePath {0}", BunkerDatabaseFilePathWeekly);
                        Console.WriteLine("ResultsFilePathWeekly {0}", ResultsFilePathWeekly);
                    }
                    if (NonWeekly)
                    {

                        string pathWrite = PathString + DatabaseFileName + @".xlsx"; //@".csv";// @"\BunkerReportDatabaseWrite.csv";
                        BunkerDatabaseFilePathMonthly = pathWrite;
                        ResultsFilePathMonthly = PathString + DatabaseFileName + @".txt";
                    }
                }
                else
                {
                    // CREATE NEW FILES
                    if (Weekly)
                    {
                        string pathWrite = PathString + CreateNewDatabase + @"Weekly" + @".xlsx";// @"\BunkerReportDatabaseWrite.csv";
                        BunkerDatabaseFilePathWeekly = pathWrite;
                        ResultsFilePathWeekly = PathString + CreateNewDatabase + @"Weekly" + @".txt";
                    }
                    if (NonWeekly)
                    {
                        string pathWrite = PathString + CreateNewDatabase + @"Monthly" + @".xlsx";// @"\BunkerReportDatabaseWrite.csv";
                        BunkerDatabaseFilePathMonthly = pathWrite;
                        ResultsFilePathMonthly = PathString + CreateNewDatabase + @"Monthly" + @".txt";
                    }
                }
            }


            if (ActionThree == true)
            {
                PathString = BunkerDatabaseDirectoryPath + @"\";
                string pathWrite = PathString + DatabaseFileName + @".xlsx";// @"\BunkerReportDatabaseWrite.csv";

                if (Weekly)
                {
                    BunkerDatabaseFilePathWeekly = pathWrite;
                    ResultsFilePathWeekly = PathString + DatabaseFileName + @".txt";
                    Console.WriteLine(BunkerDatabaseFilePathWeekly);
                }
                if (NonWeekly)
                {
                    BunkerDatabaseFilePathMonthly = pathWrite;
                    ResultsFilePathMonthly = PathString + DatabaseFileName + @".txt";
                    Console.WriteLine(BunkerDatabaseFilePathMonthly);
                }

                //Console.WriteLine(PathString);
            }

            if (!File.Exists(BunkerDatabaseFilePathWeekly) && Weekly)
            {
                //Console.WriteLine("File already exists. Oups!");

                System.IO.Directory.CreateDirectory(PathString);
                Data.State = true; // Created

                Console.WriteLine(" ================================================================");
                Console.WriteLine(" CREATE NEW EXCEL FILE: WEEKLY");
                Console.WriteLine(" ================================================================");
                Console.WriteLine();

                if (Weekly)
                {
                    //WEEKLY
                    NewHandlingTools.NewBunkerReport(BunkerDatabaseFilePathWeekly);

                    using (FileStream fs = File.Create(ResultsFilePathWeekly))
                    {
                        Byte[] info =
                            new UTF8Encoding(true).GetBytes("STATUS: Weekly Bunker Report Overview");

                        // Add some information to the file.
                        fs.Write(info, 0, info.Length);
                    }
                }
            }
            else
            { Data.State = false; }

            if (!File.Exists(BunkerDatabaseFilePathMonthly) && NonWeekly)
            {
                Console.WriteLine("File already exists. Oups!");

                System.IO.Directory.CreateDirectory(PathString);
                Data.State = true; // Created

                Console.WriteLine(" ================================================================");
                Console.WriteLine(" CREATE NEW EXCEL FILE: MONTHLY/YEARLY");
                Console.WriteLine(" ================================================================");
                Console.WriteLine();

                if (NonWeekly)
                {
                    //OVERVIEW
                    NewHandlingTools.OverviewSheet(BunkerDatabaseFilePathMonthly, CurrentFleet);


                    using (FileStream fs = File.Create(ResultsFilePathMonthly))
                    {
                        Byte[] info =
                            new UTF8Encoding(true).GetBytes("STATUS: Bunker Report Monthly/Yearly Overview");

                        // Add some information to the file.
                        fs.Write(info, 0, info.Length);
                    }
                }
            }
        }

        //================================================================
        // Exctracts The Necessary Information from the given .xlsx File
        //================================================================
        /*
        public void ExctractInfo_NEW()
        {
            Counter = 0;
            foreach (string BRFile in stringList)
            {

                Counter = Counter + 1;
                Console.WriteLine("\n File path to be examined: {0} ", BRFile);

                FileInfo File = new FileInfo(BRFile);
                //TODO ADD: Data and creation of new worksheet option
                using (var BunkerReport = new ExcelPackage(File))
                {
                    ExcelWorksheet worksheet = BunkerReport.Workbook.Worksheets[1];

                    var tempName = worksheet.Cells[6, 3].Value;

                    Console.WriteLine("\nDATA RETRIEVED:");
                    Console.WriteLine(" Vessel Name :  {0}", tempName.ToString());
                    

                    BunkerReport.Save();
                }
            }
        }
        */

        //================================================================
        // Exctracts The Necessary Information from the given .xlsx File
        //================================================================
        public void ExctractInfo()
        {
            bool eureka = false;
            Counter = 0;
            foreach (string BunkerReport in stringList)
            {
                Counter = Counter + 1;
                Console.WriteLine("\n File path to be examined: {0} ", BunkerReport);

                var excelFile = new ExcelQueryFactory(BunkerReport)//;
                {
                    ReadOnly = true
                };
                excelFile.DatabaseEngine = LinqToExcel.Domain.DatabaseEngine.Ace;
 
                /*
                var worksheetNames = excelFile.GetWorksheetNames();
                Console.WriteLine(worksheetNames);

                int i = new int();
                i = 1;
                foreach (var a0 in worksheetNames)
                {
                    if (i == 1)
                    {
                        SheetName = a0.ToString();
                        Console.WriteLine(SheetName.GetType());
                        Console.WriteLine(i);
                        Console.WriteLine("Sheetname: {0}", SheetName);
                    }
                    i = i + 1;
                }
                */
                // TODO: BECAUSE THE FILE IN LOCKED THE ABOVE PART OF CODE CANNOT BE USED
                // THIS INFORMATION IS HIDDEN!
                SheetName = "Report";
                Console.WriteLine("Sheetname: {0}", SheetName);

                var cellID = from c1 in excelFile.WorksheetRangeNoHeader("C6", "E6", SheetName) select c1; //Selects data within the B3 to G10 cell range
                var cellPosition = from c2 in excelFile.WorksheetRangeNoHeader("H6", "J6", SheetName) select c2;
                var cellFO = from c3 in excelFile.WorksheetRangeNoHeader("J29", "J29", SheetName) select c3;
                var cellDO = from c4 in excelFile.WorksheetRangeNoHeader("J43", "J43", SheetName) select c4;

                var cellDateTime = from c5 in excelFile.WorksheetRangeNoHeader("H7", "J7", SheetName) select c5;
                var cellRemarks = from c6 in excelFile.WorksheetRangeNoHeader("B45", "J47", SheetName) select c6;
                //var cellRemarks = from c6 in excelFile.WorksheetRangeNoHeader("B45", "B45", SheetName) select c6;

                var cellFORemaining = from c7 in excelFile.WorksheetRangeNoHeader("J26", "J26", SheetName) select c7;
                //var cellFORemaining = from c7 in excelFile.WorksheetRangeNoHeader("J26", "J26", SheetName) select c7;

                var cellDORemaining = from c8 in excelFile.WorksheetRangeNoHeader("J40", "J40", SheetName) select c8;

                foreach (var a1 in cellID)
                {
                    // Update data in VesselInfo instance
                    Vessel.Name = a1[0];
                    string VesselName = a1[0];
                }

                foreach (var a2 in cellPosition)
                {
                    Vessel.Position = a2[0];
                    string VesselPosition = a2[0];
                }

                foreach (var a3 in cellFO)
                {
                    float VesselFO = float.Parse(a3[0]);
                    Vessel.FO = VesselFO;
                }

                foreach (var a4 in cellDO)
                {
                    float VesselDO = float.Parse(a4[0]);
                    Vessel.DO = VesselDO;
                }

                foreach (var a5 in cellDateTime)
                {
                    Vessel.RecordDateTime = a5[0];
                }

                foreach (var a6 in cellRemarks)
                {
                     Vessel.Remarks = a6[0].ToString();
                }

                int ii = 0;
                foreach (var a6 in cellRemarks)
                {
                    //Console.WriteLine(" ii : {0}", ii);
                    //Console.WriteLine(" Remarks : {0}", a6[ii].ToString());
                    if (!string.IsNullOrEmpty(a6[ii].ToString()))
                    {
                        Vessel.Remarks = a6[ii].ToString();
                        //Console.WriteLine(" Vessel.Remarks : {0}", Vessel.Remarks);
                    }
                    ii = ii + 1;
                }

                foreach (var a7 in cellFORemaining)
                {
                    Vessel.FO_1 = float.Parse(a7[0]);
                }

                foreach (var a8 in cellDORemaining)
                {
                    Vessel.DO_1 = float.Parse(a8[0]);
                }

                excelFile.Dispose();

                Console.WriteLine("\nDATA RETRIEVED:");
                Console.WriteLine(" Vessel Name :  {0}", Vessel.Name);
                Console.WriteLine(" Vessel Position:  {0}", Vessel.Position);
                Console.WriteLine(" Vessel DateTime:  {0}", Vessel.RecordDateTime);
                Console.WriteLine(" FO [MT] : {0}", Vessel.FO);
                Console.WriteLine(" FO Remaing On board [MT] : {0}", Vessel.FO_1);
                Console.WriteLine(" FO [MT] : {0}", Vessel.DO);
                Console.WriteLine(" DO Remaing On board [MT] : {0}", Vessel.DO_1);
                Console.WriteLine(" Remarks : {0}", Vessel.Remarks);

                int cnt = 0;
                Console.WriteLine("EUREKA {0}", eureka);
                Console.WriteLine("cnt {0}", cnt);
                
                foreach (string line in CurrentFleet)
                {
                    Console.WriteLine("===================================================");
                    Console.WriteLine("Type of Current Fleet component {0}", line.GetType());

                    Console.WriteLine("EUREKA {0}", eureka);
                    Console.WriteLine("cnt {0}", cnt);

                    string text1 = CurrentFleet[cnt];// "String with non breaking spaces.";
                    text1 = System.Text.RegularExpressions.Regex.Replace(text1, @"\u00A0", " ");

                    string text2 = Vessel.Name;// "String with non breaking spaces.";
                    text2 = System.Text.RegularExpressions.Regex.Replace(text2, @"\u00A0", " ");

                    bool Example_v0 = String.Equals(text1, text2, StringComparison.InvariantCultureIgnoreCase);
                    Console.WriteLine("TEST {0}", Example_v0);

                    string Example_v1 = "MV \"" + text1 + "\"";
                    bool Example_v11 = String.Equals(Example_v1, text2, StringComparison.OrdinalIgnoreCase);

                    string Example_v2 = "M/V  \"" + text1 + "\"";
                    bool Example_v22 = String.Equals(Example_v1, text2, StringComparison.OrdinalIgnoreCase);

                    int Example_v3 = String.Compare(Example_v2, text2);

                    if (Example_v0 || Example_v11 || Example_v22 || Example_v3==0)
                    {
                       
                        AssetName = CurrentFleet[cnt].ToString();
                        Console.WriteLine(" AssetName {0}", AssetName);

                        eureka = true;
                        CurrentFleetStatus[cnt] = true;
                        // Use a tab to indent each line of the file.

                        Console.WriteLine("\t VESSEL IN FLEET");
                        Console.WriteLine("\t CURRENT FLEET STATUS {0}", CurrentFleetStatus[cnt]);
                        break;
                    }
                    cnt = cnt + 1;
                }

                Console.WriteLine("Exiting Function {0}", eureka);
                eureka = false;

                if (Weekly)
                {
                    int SheetID = new int();
                    SheetID = 1;
                    // WEEKLY 
                    WriteIntoDatabaseXlsx(SheetID, BunkerDatabaseFilePathWeekly, Counter);
                    Console.WriteLine(" ================================================================");
                    Console.WriteLine(" WRITE INTO EXCEL FILE: Weekly Report");
                    Console.WriteLine(" ================================================================");
                    Console.WriteLine();
                }
                if (NonWeekly)
                {
                    Console.WriteLine(" ================================================================");
                    Console.WriteLine(" WRITE INTO EXCEL FILE: Monthly/Yearly Report");
                    Console.WriteLine(" ================================================================");
                    Console.WriteLine();

                    WriteIntoDatabaseXlsxNEW(BunkerDatabaseFilePathMonthly, AssetName);
                }



            }
        }


        //================================================================
        // Writes Data into Database .xlsx and .txt Files MONTHLY
        //================================================================
        public void WriteIntoDatabaseXlsxNEW(string FilePath, string AssetName)
        {
            BunkerReportHandling NewHandlingTools = new BunkerReportHandling();

            string[] CoordinatesOfLastElement = NewHandlingTools.NewEntry(FilePath, AssetName);

            FileInfo File = new FileInfo(FilePath);
            //TODO ADD: Data and creation of new worksheet option
            using (var BunkerReport = new ExcelPackage(File))
            {
                ExcelWorksheet worksheet = BunkerReport.Workbook.Worksheets[1];

                //int colIndex_1 = worksheet.Cells[CoordinatesOfLastElement[0]].Start.Column;
                //int rowIndex_1 = worksheet.Cells[CoordinatesOfLastElement[0]].Start.Row;

                int colIndex_2 = worksheet.Cells[CoordinatesOfLastElement[1]].Start.Column;
                int rowIndex_2 = worksheet.Cells[CoordinatesOfLastElement[1]].Start.Row;

                // ADD DATA
                worksheet.Column(colIndex_2 + 1).Width = 15;
                worksheet.Column(colIndex_2 + 2).Width = 18;
                worksheet.Column(colIndex_2 + 3).Width = 18;

                string[] DateSplit = Vessel.RecordDateTime.Split(' ');

                worksheet.Cells[1, colIndex_2 + 1].Style.WrapText = true;
                int WeekNumber = new int();
                WeekNumber = colIndex_2 / 3 + 1;

                worksheet.Cells[1, colIndex_2 + 1].Value = "Week " + WeekNumber.ToString(); //+ DateSplit[0].ToString();// + "\t Difference in [MT]";
                worksheet.Cells[1, colIndex_2 + 1].Style.Font.Bold = true;
                worksheet.Cells[1, colIndex_2 + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[1, colIndex_2 + 1].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                worksheet.Cells[1, colIndex_2 + 1, 1, colIndex_2 + 3].Merge = true; //Merge columns start and end range
                worksheet.Cells[1, colIndex_2 + 1, 1, colIndex_2 + 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Alignment is center

                worksheet.Cells[rowIndex_2, colIndex_2 + 1].Value = Vessel.RecordDateTime;// DateSplit[0].ToString();
                worksheet.Cells[rowIndex_2, colIndex_2 + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells[rowIndex_2, colIndex_2 + 2].Value = Vessel.FO;
                worksheet.Cells[rowIndex_2, colIndex_2 + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells[rowIndex_2, colIndex_2 + 3].Value = Vessel.DO;
                worksheet.Cells[rowIndex_2, colIndex_2 + 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                // DATE REPORTED
                worksheet.Cells[2, colIndex_2 + 1].Value = "Reported";
                worksheet.Cells[2, colIndex_2 + 1].Style.Font.Bold = true;
                worksheet.Cells[2, colIndex_2 + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[2, colIndex_2 + 1].Style.Fill.BackgroundColor.SetColor(Color.AliceBlue);

                worksheet.Cells[2, colIndex_2 + 1, 3, colIndex_2 + 1].Merge = true; //Merge columns start and end range
                worksheet.Cells[2, colIndex_2 + 1, 3, colIndex_2 + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Alignment is center

                // FO
                worksheet.Cells[1, colIndex_2 + 2].Style.WrapText = true;
                worksheet.Cells[2, colIndex_2 + 2].Value = "FO Difference [MT]";
                worksheet.Cells[2, colIndex_2 + 2].Style.Font.Bold = true;
                worksheet.Cells[2, colIndex_2 + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[2, colIndex_2 + 2].Style.Fill.BackgroundColor.SetColor(Color.AliceBlue);

                worksheet.Cells[2, colIndex_2 + 2, 3, colIndex_2 + 2].Merge = true; //Merge columns start and end range
                worksheet.Cells[2, colIndex_2 + 2, 3, colIndex_2 + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Alignment is center

                // DO
                worksheet.Cells[1, colIndex_2 + 3].Style.WrapText = true;
                worksheet.Cells[2, colIndex_2 + 3].Value = "DO Difference [MT]";
                worksheet.Cells[2, colIndex_2 + 3].Style.Font.Bold = true;
                worksheet.Cells[2, colIndex_2 + 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[2, colIndex_2 + 3].Style.Fill.BackgroundColor.SetColor(Color.AliceBlue);

                worksheet.Cells[2, colIndex_2 + 3, 3, colIndex_2 + 3].Merge = true; //Merge columns start and end range
                worksheet.Cells[2, colIndex_2 + 3, 3, colIndex_2 + 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Alignment is center

                // Colour cells when values are out of range
                worksheet.Cells[rowIndex_2, colIndex_2 + 2].Style.Numberformat.Format = "0.00";
                if (Vessel.FO >= Math.Abs(AcceptableRangeFO) || Vessel.FO <= -Math.Abs(AcceptableRangeFO))
                {
                    worksheet.Cells[rowIndex_2, colIndex_2 + 2].Style.Font.Color.SetColor(Color.Red);
                }
                
                worksheet.Cells[rowIndex_2, colIndex_2 + 3].Style.Numberformat.Format = "0.00";
                if (Vessel.DO >= Math.Abs(AcceptableRangeDO) || Vessel.DO <= -Math.Abs(AcceptableRangeDO))
                {
                    worksheet.Cells[rowIndex_2, colIndex_2 + 3].Style.Font.Color.SetColor(Color.Red);
                }

                BunkerReport.Save();
            }
        }
        //================================================================
        // Writes Data into Database .xlsx and .txt Files WEEKLY
        //================================================================
        public void WriteIntoDatabaseXlsx(int SheetID, string FilePath, int Count)
        {
            FileInfo File = new FileInfo(FilePath);
            //TODO ADD: Data and creation of new worksheet option
            using (var BunkerReport = new ExcelPackage(File))
            {
                ExcelWorksheet worksheet = BunkerReport.Workbook.Worksheets[SheetID];

                var rowCnt = worksheet.Dimension.End.Row;
                var colCnt = worksheet.Dimension.End.Column;

                worksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Alignment is center

                Console.WriteLine("row : {0}", rowCnt);
                Console.WriteLine("col : {0}", colCnt);
              
                //Add the headers
                worksheet.Cells[rowCnt + 1, 1].Value = Counter;
                worksheet.Cells[rowCnt + 1, 2].Value = Vessel.Name;

                if (Vessel.Position == "AT PORT" || Vessel.Position == "At port" || Vessel.Position == "At Port" || Vessel.Position == "at port")
                {
                    worksheet.Cells[rowCnt + 1, 3].Value = "At Port";
                }

                if (Vessel.Position == "AT SEA" || Vessel.Position == "At sea" || Vessel.Position == "At Sea" || Vessel.Position == "at sea")
                {
                    worksheet.Cells[rowCnt + 1, 3].Value = "At Sea";
                }
                else
                {
                    worksheet.Cells[rowCnt + 1, 3].Value = "At Port";
                }

                //string[] DateSplit = Vessel.RecordDateTime.Split(' ');

                worksheet.Cells[rowCnt + 1, 4].Value = Vessel.RecordDateTime;// DateSplit[0];

                worksheet.Cells[rowCnt + 1, 5].Value = Vessel.FO;

                // Colour cells when values are out of range
                worksheet.Cells[rowCnt + 1, 5].Style.Numberformat.Format = "0.00";
                worksheet.Cells[rowCnt + 1, 6].Style.Numberformat.Format = "0.00";
                if (Vessel.FO >= Math.Abs(AcceptableRangeFO) || Vessel.FO <= -Math.Abs(AcceptableRangeFO))
                {
                    //worksheet.Cells[rowCnt + 1, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[rowCnt + 1, 5].Style.Font.Color.SetColor(Color.Red);
                }

                worksheet.Cells[rowCnt + 1, 6].Value = Vessel.FO_1;

                worksheet.Cells[rowCnt + 1, 7].Value = Vessel.DO;
                worksheet.Cells[rowCnt + 1, 7].Style.Numberformat.Format = "0.00";
                worksheet.Cells[rowCnt + 1, 8].Style.Numberformat.Format = "0.00";
                if (Vessel.DO >= Math.Abs(AcceptableRangeDO) || Vessel.DO <= -Math.Abs(AcceptableRangeDO))
                {
                    worksheet.Cells[rowCnt + 1, 7].Style.Font.Color.SetColor(Color.Red);
                }

                worksheet.Cells[rowCnt + 1, 8].Value = Vessel.DO_1;

                worksheet.Cells[rowCnt + 1, 9].Value = Vessel.Remarks;

                BunkerReport.Save();
            }
        }

        //================================================================
        // From the Fleet.txt reads the names of the existing ships
        //================================================================
        public string[] ReadTxt()
        {
            // Example #2
            // Read each line of the file into a string array. Each element
            // of the array is one line of the file.

            string path = System.Environment.CurrentDirectory;
            path = path + @"\Fleet.txt";
            Console.WriteLine("GET CURRENT DIRECTORY {0}", path);
            string[] lines = System.IO.File.ReadAllLines(path);
            
            CurrentFleet = lines;

            int rowsOrHeight = CurrentFleet.GetLength(0);
            CurrentFleetStatus = new bool[rowsOrHeight];

            Console.WriteLine("Array size rows {0}", rowsOrHeight);

            // Display the file contents by using a foreach loop.
            System.Console.WriteLine("Contents of Fleet.txt = ");
            foreach (string line in lines)
            {
                // Use a tab to indent each line of the file.
                Console.WriteLine("\t" + line);
            }
            return lines;
        }

        //================================================================
        // APPLICATION ACTIONS AND REACTIONS
        // Button 1: Bunker Report Directory Activated
        // Button 3: Bunker Database Directory Activated
        // Button 2: All Data complete and proceed
        // Action 2 => New Database in .xlsx
        // Action 3 => Update Old Database in .xlsx
        //================================================================

        // SELECT SOURCE FOLDER
        private void button1_Click(object sender, EventArgs e)
        {
            ActionOne = true;

            CreateNewDatabase = DateTime.Today.ToString("dd_MM_yyyy");
            Console.WriteLine(" ================================================================");
            Console.WriteLine(" Today {0}", CreateNewDatabase);
            Console.WriteLine(" BUTTON 1 ");
            Console.WriteLine(" Action 1: true");
            Console.WriteLine(" ================================================================");
            Console.WriteLine();
            Console.WriteLine(" Call to ChooseFolderReports...");

            ChooseFolderReports();
            Console.WriteLine();
            Console.WriteLine("\t" + "BunkerReportFilePath: {0} ", BunkerReportDirectoryPath);
        }

        // SELECT DATABASE DIRECTORY
        private void button3_Click(object sender, EventArgs e)
        {
            Console.WriteLine(" ================================================================");
            Console.WriteLine(" BUTTON 3 ");
            Console.WriteLine(" ================================================================");
            Console.WriteLine();
            Console.WriteLine(" Call to ChooseFolderDatabase...");
            ChooseFolderDatabase();
            Console.WriteLine();
            Console.WriteLine("\t" + "BunkerDatabaseFilePath: {0} ", BunkerDatabaseDirectoryPath);
        }

        

        // PROCEED
        private void button2_Click(object sender, EventArgs e)
        {
            Console.WriteLine(" ================================================================");
            Console.WriteLine(" BUTTON 2 ");
            Console.WriteLine(" ================================================================");
            Console.WriteLine();
            Console.WriteLine(" ========================================================================");
            Console.WriteLine(" INITIALIZING: BUNKER REPORT DATABASE APPLICATION");
            Console.WriteLine(" ========================================================================");

            if (checkBox1.Checked)
            {
                Weekly = true;
            }
            else
            {
                Weekly = false;
            }

            if (checkBox4.Checked)
            {
                NonWeekly = true;
            }
            else
            {
                NonWeekly = false;
            }
            //============================================================================================
            // Directory: BUNKER REPORTS
            //============================================================================================
            if ((ActionOne == true) && (!string.IsNullOrWhiteSpace(textBox1.Text))) // If button 1 has been clicked
            {
                Console.WriteLine("Action 1 = true");
                // In case it is changed by the user, save the altered path!
                BunkerReportDirectoryPath = textBox1.Text;
                Console.WriteLine("Action 1. BunkerReportFilePath: {0}", BunkerReportDirectoryPath);

                Console.WriteLine("Path from text box 1 {0}", textBox1.Text);
                Console.WriteLine("BunkerReportFilePath {0}", BunkerReportDirectoryPath);
                Console.WriteLine();

                //AcceptableBunkerDirectoryPath = true;

                FindBunkerReportFilePaths();

            }
            else if (ActionOne == false && (!string.IsNullOrWhiteSpace(textBox1.Text)))
            {
                BunkerReportDirectoryPath = textBox1.Text;
                Console.WriteLine("From TextBox1 BunkerReportFilePath: {0}", BunkerReportDirectoryPath);
                Console.WriteLine();

                //AcceptableBunkerDirectoryPath = true;

                FindBunkerReportFilePaths();
            }
            else if (!string.IsNullOrWhiteSpace(textBox1.Text))
            {
                BunkerReportDirectoryPath = textBox1.Text;
                Console.WriteLine("From TextBox1 BunkerReportFilePath: {0}", BunkerReportDirectoryPath);
                Console.WriteLine();

                FindBunkerReportFilePaths();
            }
            else
            { 
                MessageBox.Show("Please Complete the fields above!");
            }

            //============================================================================================
            // Path File: BUNKER DATABASE
            //============================================================================================
            // Where to put the final excel with the exctracted data? huh?
            if (!string.IsNullOrWhiteSpace(textBox2.Text))
            {
                // Create New Database
                BunkerDatabaseDirectoryPath = textBox2.Text;
                Console.WriteLine("From TextBox2 BunkerDatabaseDirectoryPath: {0}", BunkerDatabaseDirectoryPath);
                AcceptableDatabaseDirectoryPath = true;
            }

            // Check Button4
            if (!string.IsNullOrWhiteSpace(textBox4.Text))
            {
                DatabaseFileName = textBox4.Text;
                DatabaseFileNameState = true;
            }
            else
            {
                DatabaseFileNameState = false;
            }

            //==============================================================================================
            // Make sure the form is completed correctly
            //==============================================================================================
            if (ActionTwo || ActionThree)
            {
                AcceptableChoiceNewUpdate = true;
            }
            else if (!ActionTwo && !ActionThree) // If bunker report folder has been selected
            {
                MessageBox.Show("Please Select 'New' or 'Update' to continue");
            }

            Console.WriteLine("AcceptableChoiceNewUpdate {0}", AcceptableChoiceNewUpdate);
            Console.WriteLine("AcceptableBunkerDirectoryPath {0}", AcceptableBunkerDirectoryPath);
            Console.WriteLine("AcceptableDatabaseDirectoryPath {0}", AcceptableDatabaseDirectoryPath);

            
            if (AcceptableChoiceNewUpdate && AcceptableBunkerDirectoryPath && AcceptableDatabaseDirectoryPath && (Weekly || NonWeekly))
            {
                

                progressBar1.Value = 20;

                if (UserDefinedRangeFO)
                {
                    // Set conditional Formatting values
                    AcceptableRangeFO = float.Parse(textBox3.Text);  // [metric tonnes]
                }
                if (UserDefinedRangeDO)
                {
                    AcceptableRangeDO = float.Parse(textBox5.Text);   // [metric tonnes]
                }

                Console.WriteLine(" Call to FindBunkerReportFilePaths...");
                Console.WriteLine();
                FindBunkerReportFilePaths();

                progressBar1.Value = 40;

                Console.WriteLine(" Call to ReadTxt...");
                Console.WriteLine();
                CurrentFleet = ReadTxt();

                progressBar1.Value = 50;

                Console.WriteLine(" Bunker Report Folden Path has been selected");
                Console.WriteLine();

                Console.WriteLine(" Call to DatabaseHandling...");
                Console.WriteLine();

                DataBaseHandling();

                progressBar1.Value = 70;
                Console.WriteLine(" Database .xlsx file created status: {0}", Data.State);

                Console.WriteLine(" Call to ExtractInfo...");
                Console.WriteLine();
                //ExctractInfo_NEW();

                ExctractInfo();
                progressBar1.Value = 80;
                Console.WriteLine("\nDATA RETRIEVED:");
                Console.WriteLine(" Vessel Name :  {0}", Vessel.Name);
                Console.WriteLine(" Vessel Position:  {0}", Vessel.Position);
                Console.WriteLine(" Vessel DateTime:  {0}", Vessel.RecordDateTime);
                Console.WriteLine(" FO [MT] : {0}", Vessel.FO);
                Console.WriteLine(" FO Remaing On board [MT] : {0}", Vessel.FO_1);
                Console.WriteLine(" FO [MT] : {0}", Vessel.DO);
                Console.WriteLine(" DO Remaing On board [MT] : {0}", Vessel.DO_1);
                Console.WriteLine(" Remarks : {0}", Vessel.Remarks);

                string FleetPath = System.Environment.CurrentDirectory;
                FleetPath = FleetPath + @"\Fleet.txt";
                Console.WriteLine("GET CURRENT DIRECTORY {0}", FleetPath);

                if (Weekly)
                {
                    Console.WriteLine(" ResultsFilePathWeekly: {0}", ResultsFilePathWeekly);

                    

                    // OUTPUT
                    using (System.IO.StreamWriter file =
                                new System.IO.StreamWriter(ResultsFilePathWeekly, true))
                    {
                        file.WriteLine("");
                        file.WriteLine("========================================================");
                        file.WriteLine(" Last Update at: {0}", CreateNewDatabase);
                        file.WriteLine(" Bunker Reports PENDING for the following VESSELS:");
                        file.WriteLine("");
                        file.WriteLine(" To update vessel names information in Fleet.txt");
                        file.WriteLine(" PATH: {0}", FleetPath);
                        file.WriteLine("========================================================");
                    }
                }
                if (NonWeekly)
                {
                    Console.WriteLine(" ResultsFilePathOverview: {0}", ResultsFilePathMonthly);

                    // OUTPUT
                    using (System.IO.StreamWriter file =
                                new System.IO.StreamWriter(ResultsFilePathMonthly, true))
                    {
                        file.WriteLine("");
                        file.WriteLine("========================================================");
                        file.WriteLine(" Last Update at: {0}", CreateNewDatabase);
                        file.WriteLine(" Bunker Reports PENDING for the following VESSELS:");
                        file.WriteLine("");
                        file.WriteLine(" To update vessel names information in Fleet.txt");
                        file.WriteLine(" PATH: {0}", FleetPath);
                        file.WriteLine("========================================================");
                    }
                }

                progressBar1.Value = 90;
                int cnt = 0;
                foreach (bool ship in CurrentFleetStatus)
                {
                    if (CurrentFleetStatus[cnt] == true)
                    { }
                    else
                    {
                        if (Weekly)
                        {
                            using (System.IO.StreamWriter file =
                                new System.IO.StreamWriter(ResultsFilePathWeekly, true))
                            {
                                file.WriteLine(CurrentFleet[cnt]);
                            }
                        }
                        if (NonWeekly)
                        {
                            using (System.IO.StreamWriter file =
                                new System.IO.StreamWriter(ResultsFilePathMonthly, true))
                            {
                                file.WriteLine(CurrentFleet[cnt]);
                            }
                        }
                    }

                    cnt = cnt + 1;
                }
                
                progressBar1.Value = 100;
                MessageBox.Show(" Succesfull Bunker Report Database Update", "Status");
            }
            else
            {
                MessageBox.Show("Please Complete the fields above!");
            }
        }


        private void label1_Click_1(object sender, EventArgs e)
        {
            
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {   // create newdatabase
            ActionTwo = true;
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {   
            // update newdatabase
            ActionThree = true;

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            UserDefinedRangeFO = true;
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            UserDefinedRangeDO = true;
        }

        private void progressBar1_Click(object sender, EventArgs e)
        {
            progressBar1.Maximum = 100;
            progressBar1.Step = 1;
            progressBar1.Value = 0;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            //Weekly = true;
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            //NonWeekly = true;
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}
