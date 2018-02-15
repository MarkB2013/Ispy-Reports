using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ISpy_Reports
{
    public partial class Form1 : Form
    {
        //Variable declaration for DataTable that will output dex file meter readings to excel sheet
        DataTable Data = new DataTable();
        DataSet dataSet = new DataSet();
        DataRow row;

        //Directory variable declarations
        static DirectoryInfo DexFiles = new DirectoryInfo(Properties.Settings.Default.Workbench_Directory + Properties.Settings.Default.DexFilePath);
        static DirectoryInfo ExcelFile = new DirectoryInfo(Properties.Settings.Default.Workbench_Directory + Properties.Settings.Default.ExcelFilePath);
        static DirectoryInfo Output = new DirectoryInfo(Properties.Settings.Default.Workbench_Directory + Properties.Settings.Default.Output);
        static DirectoryInfo ErrorLog = new DirectoryInfo(Properties.Settings.Default.Workbench_Directory + Properties.Settings.Default.ErrorLog);
        static DirectoryInfo MachineArchive = new DirectoryInfo(Properties.Settings.Default.Workbench_Directory + Properties.Settings.Default.MachineDatabase);
        static DirectoryInfo OldReports = new DirectoryInfo(Properties.Settings.Default.Workbench_Directory + Properties.Settings.Default.OldReports);

        //List that will contain any errors that occur during runtime
        List<string> masterSessionFailureLog = new List<string>();

        string currentDataFile = "";
        string currentMachineNumber = "";
        string currentExcelMachine = "";
        string currentRecordDate = "";

        public Form1()
        {
            //String builder variables for outputing console data to log file
            StringBuilder sb = new StringBuilder();
            StringWriter sw = new StringWriter(sb);

            //Set console to output to string writer
            Console.SetOut(sw);

            InitializeComponent();

            //Check if Excel file exists before continuing
            if (File.Exists(ExcelFile.ToString()))
            {
                Console.WriteLine("Initializing...");
                Console.WriteLine("Performing File Operations...");

                try
                {
                    //BuildReport matches the dex file meter readings with the machine data from an input Excel file, then uses ExportDataSet method to export data to Excel spreadsheet
                    BuildReport();
                }
                catch(Exception sessionFailure)
                {
                    //Add any errors to log file list
                    masterSessionFailureLog.Add(sessionFailure.Message);
                }                
            }
            else
            {
                //Else if input Excel file, add message to log file list
                masterSessionFailureLog.Add("Machine data excel file not found, Export EasiTrax User Report 6 as CSV file to: " + ExcelFile.ToString());
                Console.WriteLine("Machine data excel file not found, Export EasiTrax User Report 6 as CSV file to: " + ExcelFile.ToString());
            }

            //Get Excel file age
            DateTime excelDate = File.GetLastWriteTime(ExcelFile.ToString());
            TimeSpan excelAge = DateTime.Now - excelDate;
            
            //IF Excel file age is more than 12 hours, add message to log file list
            if(excelAge.TotalHours > 12)
            {
                masterSessionFailureLog.Add("Machine data excel file is out of date, Export EasiTrax User Report 6 as CSV file to: " + ExcelFile.ToString());
                Console.WriteLine("Machine data excel file is out of date, Export EasiTrax User Report 6 as CSV file to: " + ExcelFile.ToString());
            }

            //If any errors occur, write errors to log file
            if (masterSessionFailureLog.Count > 0)
            {
                File.WriteAllLines(ErrorLog + @"Master_Log\" + "Master_Error_Log_" + DateTime.Now.ToString("dd-MMM-yyyy_HH-mm-ss") + ".txt", masterSessionFailureLog);
            }

            //String builder list
            List<string> sbList = new List<string>();

            //Build string from string writer console output
            sw.Close();
            StringReader sr = new StringReader(sb.ToString());
            string completeString = sr.ReadToEnd();
            sr.Close();

            //Add console output to list
            sbList.Add(sb.ToString());

            //Write console output to log file folder
            File.WriteAllLines(ErrorLog + @"Command Prompt Text Logs\" + "Command_Log_" + DateTime.Now.ToString("dd-MMM-yyyy_HH-mm-ss") + ".txt", sbList);

            //End program
            Environment.Exit(0000);
        }        

        //Export Method, template from DocumentFormat NuGet package
        private void ExportDataSet(DataSet ds, string destination)
        {
            try
            {
                using (var workbook = SpreadsheetDocument.Create(destination, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
                {
                    var workbookPart = workbook.AddWorkbookPart();

                    workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook
                    {
                        Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets()
                    };

                    foreach (System.Data.DataTable table in ds.Tables)
                    {
                        var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                        var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
                        sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                        DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                        string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                        uint sheetId = 1;
                        if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
                        {
                            sheetId =
                                sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                        }

                        DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
                        sheets.Append(sheet);

                        DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

                        List<String> columns = new List<string>();
                        foreach (System.Data.DataColumn column in table.Columns)
                        {
                            columns.Add(column.ColumnName);

                            DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell
                            {
                                DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String,
                                CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName)
                            };
                            headerRow.AppendChild(cell);
                        }

                        sheetData.AppendChild(headerRow);

                        foreach (System.Data.DataRow dsrow in table.Rows)
                        {
                            DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                            foreach (String col in columns)
                            {
                                DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell
                                {
                                    DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String,
                                    CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()) //
                                };
                                newRow.AppendChild(cell);
                            }

                            sheetData.AppendChild(newRow);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Export failed, please make sure export file is not in open in another program and try again. " + e);
            }
        }

        //Defualt load properties method
        private void Load_Properties()
        {
            try
            {
                //Configuration file name
                string configFile = "config.cfg";

                //Get current directory for program and combine with name to get full directory path for file
                string path = Path.Combine(Environment.CurrentDirectory, @"Data\", configFile);
            }

            catch (Exception e)
            {
                MessageBox.Show("Failed to load properties file, please try again. if this problem persists, re-install the program." + e);
            }
        }

        private void BuildReport()
        {
            //List for machine archive file names
            List<string> datFileNames = new List<string>();

            //List for machine database dex files
            List<string> DexNames = new List<string>();

            //List for machine vend meters
            List<int> MachineVendCounts = new List<int>();

            //Error log list
            List<string> Error = new List<string>();

            //Read all lines of input Excel file
            string[] ExcelLines = File.ReadAllLines(ExcelFile.ToString());

            //Get count of lines in Excel input file
            int ExcelCount = ExcelLines.Length;

            //Set datetime variables of previous 7 days
            DateTime Today = DateTime.Today;
            var yesterday = Today.AddDays(-1);
            var twoDays = Today.AddDays(-2);
            var threeDays = Today.AddDays(-3);
            var fourDays = Today.AddDays(-4);
            var fiveDays = Today.AddDays(-5);
            var sixDays = Today.AddDays(-6);
            var sevenDays = Today.AddDays(-7);
            var eightDays = Today.AddDays(-8);

            #region Report Columns

            //Add columns to the DataTable

            //Machine number column details
            DataColumn machineNumber = new DataColumn
            {
                DataType = Type.GetType("System.Int32"),
                ColumnName = "Machine Number",
                ReadOnly = false,
                Unique = true,
                AutoIncrement = false
            };           

            //Customer column details
            DataColumn customer = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Customer",
                AutoIncrement = false
            };         

            //Total vend count from MEI EasiTrax column details
            DataColumn meiTotalVendCount = new DataColumn
            {
                DataType = Type.GetType("System.Int32"),
                ColumnName = "MEI Total Vend Count",
                AutoIncrement = false
            };

            //DEX file vend count column details
            DataColumn dexVendCount = new DataColumn
            {
                DataType = Type.GetType("System.Int32"),
                ColumnName = "DEX Total Vend Count",
                AutoIncrement = false
            };

            //Number of stock sold column details
            DataColumn stocksold = new DataColumn
            {
                DataType = Type.GetType("System.Int32"),
                ColumnName = "Stock Sold",
                AutoIncrement = false
            };

            //Capacity column details
            DataColumn capacity = new DataColumn
            {
                DataType = Type.GetType("System.Int32"),
                ColumnName = "Capacity",
                AutoIncrement = false
            };

            //Current machine stock level column details
            DataColumn currentStock = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Current Sold Stock %",
                AutoIncrement = false
            };

            //Date of DEX file column details
            DataColumn dateTimeDex = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Date/Time of DEX",
                AutoIncrement = false
            };
            

            //Predicted stock level on day of next scheduled driver visit column details
            DataColumn visitStockPrediction = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Next Visit Stock Prediction",
                AutoIncrement = false
            };

            //Next shceduled driver visit column details
            DataColumn nextVisit = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Next Scheduled Visit Date",
                AutoIncrement = false
            };

            //Predicted date of when machine will have sold 30% of stock column details
            DataColumn OptimalFill = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Optimal Fill Date",
                AutoIncrement = false
            };

            //Days to next scheduled driver visit column details
            DataColumn visitIn = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Scheduled Visit In:",
                AutoIncrement = false
            };

            //Predicted days until machine has sold 30% of its stock column details
            DataColumn OptimalFillIn = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Optimal Fill Date In:",
                AutoIncrement = false
            };

            //Day of the week of next scheduled driver visit column details
            DataColumn visitDay = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Scheduled Visit Day",
                AutoIncrement = false
            };

            //Day of week of predicted date when machine has sold 30% of its stock column details
            DataColumn OptimalFillDay = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Optimal Fill Day",
                AutoIncrement = false
            };

            //Days until machine has sold 40% of stock column details
            DataColumn Machine40 = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Machine @40% Stock Sold in",
                AutoIncrement = false
            };

            //Days until machine has sold 30% stock column details
            DataColumn Machine30 = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Machine @30% Stock Sold in",
                AutoIncrement = false
            };

            //Route number for machine column details
            DataColumn routeNumber = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Route Number",
                AutoIncrement = false
            };

            //Name of driver on route for machine column details
            DataColumn routeDriverName = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Route Driver Name",
                AutoIncrement = false
            };

            //Telemetry provider column details
            DataColumn telemetryProvider = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Telemetry Provider",
                AutoIncrement = false
            };

            //Telemetry ID column details
            DataColumn telemetryID = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Telemetry ID",
                AutoIncrement = false
            };

            //Count of days since last visit column details
            DataColumn daysSinceLastFill = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Days since last refill",
                AutoIncrement = false
            };


            //Machine's sector column details
            DataColumn sector = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Sector",
                AutoIncrement = false
            };

            //Machine products column details
            DataColumn products = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Products",
                AutoIncrement = false
            };

            //Type of machine column details
            DataColumn machineType = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Machine Type",
                AutoIncrement = false
            };

            //Avgerage sales calculated from last 7 days column details
            DataColumn AvgWeekSales = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Average Sales Per Week",
                AutoIncrement = false
            };

            //Meter readings from previous day 1 column
            DataColumn PreviousDay1Sales = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = yesterday.DayOfWeek.ToString() + "|" + yesterday.ToString("dd-MMM-yy"),
                AutoIncrement = false
            };

            //Meter readings from previous day 2 column
            DataColumn PreviousDay2Sales = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = twoDays.DayOfWeek.ToString() + "|" + twoDays.ToString("dd-MMM-yy"),
                AutoIncrement = false
            };

            //Meter readings from previous day 3 column
            DataColumn PreviousDay3Sales = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = threeDays.DayOfWeek.ToString() + "|" + threeDays.ToString("dd-MMM-yy"),
                AutoIncrement = false
            };

            //Meter readings from previous day 4 column
            DataColumn PreviousDay4Sales = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = fourDays.DayOfWeek.ToString() + "|" + fourDays.ToString("dd-MMM-yy"),
                AutoIncrement = false
            };

            //Meter readings from previous day 5 column
            DataColumn PreviousDay5Sales = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = fiveDays.DayOfWeek.ToString() + "|" + fiveDays.ToString("dd-MMM-yy"),
                AutoIncrement = false
            };

            //Meter readings from previous day 6 column
            DataColumn PreviousDay6Sales = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = sixDays.DayOfWeek.ToString() + "|" + sixDays.ToString("dd-MMM-yy"),
                AutoIncrement = false
            };

            //Meter readings from previous day 7 column
            DataColumn PreviousDay7Sales = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = sevenDays.DayOfWeek.ToString() + "|" + sevenDays.ToString("dd-MMM-yy"),
                AutoIncrement = false
            };

            DataColumn lastVisitDate = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Last Visit Date",
                AutoIncrement = false
            };

            //Add columns
            Data.Columns.Add(machineNumber);           
            Data.Columns.Add(customer);

            Data.Columns.Add(meiTotalVendCount);
            Data.Columns.Add(dexVendCount);

            Data.Columns.Add(dateTimeDex);
            Data.Columns.Add(daysSinceLastFill);
            Data.Columns.Add(lastVisitDate);

            Data.Columns.Add(capacity);
            Data.Columns.Add(stocksold);
            Data.Columns.Add(currentStock);

            Data.Columns.Add(nextVisit);
            Data.Columns.Add(visitIn);
            Data.Columns.Add(visitDay);
            Data.Columns.Add(visitStockPrediction);

            Data.Columns.Add(OptimalFill);
            Data.Columns.Add(OptimalFillIn);
            Data.Columns.Add(OptimalFillDay);

            Data.Columns.Add(Machine40);
            Data.Columns.Add(Machine30);

            Data.Columns.Add(routeNumber);
            Data.Columns.Add(routeDriverName);

            Data.Columns.Add(telemetryProvider);
            Data.Columns.Add(telemetryID);

            Data.Columns.Add(sector);
            Data.Columns.Add(products);
            Data.Columns.Add(machineType);

            Data.Columns.Add(AvgWeekSales);

            Data.Columns.Add(PreviousDay1Sales);
            Data.Columns.Add(PreviousDay2Sales);
            Data.Columns.Add(PreviousDay3Sales);
            Data.Columns.Add(PreviousDay4Sales);
            Data.Columns.Add(PreviousDay5Sales);
            Data.Columns.Add(PreviousDay6Sales);
            Data.Columns.Add(PreviousDay7Sales);

            #endregion

            //Add file names of each dex file in DexFiles directory to list
            foreach (FileInfo DexFile in DexFiles.GetFiles("*.dex"))
            {
                DexNames.Add(DexFile.Name);
            }

            //Add file names of each machine archive file in MachineArchive directory to list
            foreach (FileInfo file in MachineArchive.GetFiles("*.dat"))
            {
                datFileNames.Add(file.Name);
            }

            //Simple count for displaying file progression in console
            int simpleCount = 0;


            //Loop through each line in excel file
            foreach (string Line in ExcelLines)
            {
                try
                {
                    //New data row for each line in Excel file
                    row = Data.NewRow();

                    //Add 1 to simple counter
                    simpleCount++;

                    //Display current file count
                    Console.WriteLine(simpleCount + " out of " + ExcelCount);

                    //Split data in line by delimiter
                    string[] LineData = Line.Split(',');

                    //Excel machine data variable declarations

                    //EasiTrax machine number
                    string machinenumber = LineData[0].Trim('"');

                    //Location of machine
                    string machineLocation = LineData[2].Trim('"');

                    //Provider of telemetry unit in machine
                    string TelemetryProvider = LineData[8].Trim('"');

                    //Capacity of machine
                    int Capacity = Int32.Parse(LineData[10].Trim('"'));

                    //Vend meter reading from EasiTrax
                    int MEIVendCount = Int32.Parse(LineData[11].Trim('"'));

                    //Cash meter reading from EasiTrax
                    string MEICashCount = LineData[12].Trim('"');

                    //Parse date/time of last driver visit
                    DateTime LastVisitDate = DateTime.ParseExact(LineData[13].Trim('"'), "ddMMyy", CultureInfo.InstalledUICulture);

                    //PHYSID of telemetry unit
                    string machinePHYSID = LineData[14].Trim('"');

                    //Parse date/time of next scheduled driver visit
                    DateTime NextScheduledVisit = DateTime.ParseExact(LineData[15].Trim('"'), "ddMMyy", CultureInfo.InstalledUICulture);

                    NextScheduledVisit = NextScheduledVisit.AddHours(16);

                    //Route number of scheduled driver
                    string RouteName = LineData[17].Trim('"');

                    //Scheduled driver name
                    string DriverName = LineData[18].Trim('"');

                    //Model of machine
                    string machineModel = LineData[19].Trim('"');

                    //Type of machine
                    string MachineType = LineData[20].Trim('"');

                    //Machine location type
                    string MachineSector = LineData[21].Trim('"');

                    //Days until next scheduled visit timespan                  
                    TimeSpan ScheduleSplit = NextScheduledVisit - DateTime.Now;

                    string nextVisitToString = ScheduleSplit.TotalDays.ToString("0.0") + " Days";

                    if (ScheduleSplit.TotalDays < 1)
                    {
                        nextVisitToString = "Today";
                    }

                    

                    //Set machine number column variable to machine number value  
                    row["Machine Number"] = machinenumber;

                    //Set machine location column to machine location  variable value 
                    row["Customer"] = machineLocation;

                    //Set machine capacity column to capacity  variable value 
                    row["Capacity"] = Capacity;

                    //Set next scheduled visit date column to next scheduled visit parsed date variable value 
                    row["Next Scheduled Visit Date"] = NextScheduledVisit.ToString("dd MMM yy");

                    row["Last Visit Date"] = LastVisitDate.ToString("dd MMM yy");

                    //Set days to next schedule visit column to calculated next schedule visit date variable value
                    row["Scheduled Visit In:"] = nextVisitToString;

                    //Set day of week of next scheduled visit column to next scheduled visit day of week variable value
                    row["Scheduled Visit Day"] = NextScheduledVisit.DayOfWeek;

                    //Set route number column to route number variable
                    row["Route Number"] = RouteName;

                    //Set driver name column to driver name variable value
                    row["Route Driver Name"] = DriverName;

                    //Set telemetry provider column to telemetry provider variable value
                    row["Telemetry Provider"] = TelemetryProvider;

                    //Set PHYSID column to PHYSID variable value
                    row["Telemetry ID"] = machinePHYSID;

                    //Split date time variable to only get counted hours
                    string DateTimeSplit = ((DateTime.Now - LastVisitDate).TotalDays).ToString("0.0");

                    //Set days since last refill column to calculated days since last visit variable value
                    row["Days since last refill"] = DateTimeSplit + " Days";

                    //Set machine sector column to machine sector vairable value
                    row["Sector"] = MachineSector;

                    //Set machine type column to machine type variable value
                    row["Products"] = MachineType;

                    //Set machine type column to machine type variable value
                    row["Machine Type"] = machineModel;

                    //try
                    //{
                        //foreach loop for each file in dex file names list
                        foreach (string FileName in DexNames)
                        {
                            //Split file name by delimiter to array
                            string[] DexNameData = FileName.Split('_', '.', '-');

                            //Set PHYSID variable from file name value
                            string DexPHYSID = DexNameData[0];

                            //Set date of dex file from last modified time in dex file
                            DateTime DexDate = File.GetLastWriteTime(DexFiles + FileName);

                            //Variables for later use in retreiving data from multiple lines in the dex file
                            int OldestReading = 0;
                            int NewestReading = 0;

                            //If PHYSID from Excel input file matches Dex PHYSID
                            if (machinePHYSID == DexPHYSID)
                            {
                                //Full meters variable declaration
                                string MeterLine = "";

                                currentMachineNumber = machinenumber;

                                //Path to current dex file
                                string DexFilePath = DexFiles.ToString() + FileName;

                                //Read all lines in dex file and add to an array
                                string[] DexLines = File.ReadAllLines(DexFilePath);

                                //Foreach loop for every line in dex
                                foreach (string DexLine in DexLines)
                                {
                                    //Using Array.Find to find meter line and assign the value to the meter line variable
                                    MeterLine = Array.Find(DexLines,
                                element => element.StartsWith("VA1", StringComparison.Ordinal));
                                }

                                //Dex meter variable declaration
                                int DexCashMeterValue = 0;
                                int DexVendMeterValue = 0;

                                //If the meter line was found
                                if (MeterLine != null)
                                {
                                    //Split meter line by delimiter
                                    string[] MeterReads = MeterLine.Split('*');

                                    //Assign split values to relevant variables
                                    DexCashMeterValue = Int32.Parse(MeterReads[1]);
                                    DexVendMeterValue = Int32.Parse(MeterReads[2]);
                                }

                                //Percent sold calculations
                                decimal PercentSold = 0;
                                decimal StockSold = DexVendMeterValue - MEIVendCount;
                                decimal StockLeft = Capacity - StockSold;

                                //Set EasiTrax meter reading column to meter reading variable value
                                row["MEI Total Vend Count"] = MEIVendCount;

                                //If the machine has sales
                                if (StockSold != 0)
                                {
                                    //Calulate the percentage by dividng the total sales by the capacity of machine and multiply it by 100
                                    PercentSold = (StockSold / Capacity) * 100;
                                }
                                else
                                {
                                    //If there are no sales then percent sold is 0
                                    PercentSold = 0;
                                }

                                //Decimal to int parsing to remove unwanted extra values
                                int IntPercent = Int32.Parse(PercentSold.ToString("0"));

                                //Set Dex date column to Dex date variable value
                                row["Date/Time of DEX"] = DexDate.ToString("dd MMM yy HH:mm:ss");

                                //Set total Dex vend count column to Dex vend meter variable value
                                row["DEX Total Vend Count"] = DexVendMeterValue;

                                //Set PHYSID column to PHYSID variable value
                                row["Telemetry ID"] = machinePHYSID;

                                //If percent sold is within range set current stock column to percent sold variable value, else set it to bad reading to indicate meter reading error in Excel report
                                if (PercentSold > 0 && PercentSold < 100)
                                {
                                    row["Current Sold Stock %"] = IntPercent + " %";
                                }
                                else
                                {
                                    row["Current Sold Stock %"] = "Bad Reading";
                                }

                                //If stock sold is within expected range set stock sold column to vend count meter reading variable value, else set it to -1 value to indicate meter reading error
                                if (StockSold > 0 && StockSold < Capacity)
                                {
                                    row["Stock Sold"] = StockSold;
                                }
                                else
                                {
                                    row["Stock Sold"] = -1;
                                }

                                //Day count variable declarations
                                int dayEight = 0;
                                int daySeven = 0;
                                int daySix = 0;
                                int dayFive = 0;
                                int dayFour = 0;
                                int dayThree = 0;
                                int dayTwo = 0;
                                int dayOne = 0;
                                int WeekAverage = 0;

                                //DateTime variables used for finding latest record in machine archive file
                                DateTime oldest = DateTime.Now;
                                DateTime newest = DateTime.Now.AddDays(-35);

                                //Foreach loop for every data file in machine archive directory
                                foreach (string datFile in datFileNames)
                                {
                                    //Read all lines from file and add to an array
                                    string[] datFileData = File.ReadAllLines(MachineArchive + datFile);

                                    currentDataFile = datFile;
                                    //try
                                    //{
                                        //Foreach loop for each line in file
                                        foreach (string datLine in datFileData)
                                        {
                                            //Record date/time declaration set to current as default
                                            DateTime RecordDate = DateTime.Now;

                                            //Machine vend meter variable declaration
                                            int MachineTotalVendMeter = 0;

                                            //Line and entry splitting into array to isolate data variables
                                            string[] datLineData = datLine.Split('_');
                                            string[] datEntries = datLineData[0].Split('*');
                                            string[] EntryMachineNumberData = datEntries[1].Split('-');
                                            string[] EntryDateData = datEntries[0].Split('-');

                                            //Parse record date from date/time entry
                                            RecordDate = DateTime.ParseExact(EntryDateData[1], "ddMMyyyy,HHmmss", CultureInfo.InstalledUICulture);

                                            currentRecordDate = RecordDate.ToString("dd/MMM/yy HH:mm:ss");

                                            //Set entry machine number to variable
                                            string EntryMachineNumber = EntryMachineNumberData[1];

                                            

                                            //Check if machine number in data file matches machine number of dex file
                                            if (EntryMachineNumber == machinenumber)
                                            {
                                                //Split entry data from entry tag
                                                string[] VendMeterData = datEntries[12].Split('-');

                                                //Set machine vend count variable to parsed value from array
                                                MachineTotalVendMeter = Int32.Parse(VendMeterData[1]);

                                                //If record date is older than previous record set oldest date/time variable to current entry date/time
                                                if (RecordDate < oldest)
                                                {
                                                    OldestReading = MachineTotalVendMeter;
                                                    oldest = RecordDate;
                                                }

                                                //If record date is newer than previous record set newest date/time variable to current entry date/time
                                                if (RecordDate > newest)
                                                {
                                                    NewestReading = MachineTotalVendMeter;
                                                    newest = RecordDate;
                                                }

                                                //Check if entry matches desired date and applies value to variable if it does
                                                if (RecordDate.ToString("ddMMyy") == Today.ToString("ddMMyy"))
                                                {
                                                    dayOne = MachineTotalVendMeter;
                                                }
                                                if (RecordDate.ToString("ddMMyy") == twoDays.ToString("ddMMyy"))
                                                {
                                                    dayTwo = MachineTotalVendMeter;
                                                }
                                                if (RecordDate.ToString("ddMMyy") == threeDays.ToString("ddMMyy"))
                                                {
                                                    dayThree = MachineTotalVendMeter;
                                                }
                                                if (RecordDate.ToString("ddMMyy") == fourDays.ToString("ddMMyy"))
                                                {
                                                    dayFour = MachineTotalVendMeter;
                                                }
                                                if (RecordDate.ToString("ddMMyy") == fiveDays.ToString("ddMMyy"))
                                                {
                                                    dayFive = MachineTotalVendMeter;
                                                }
                                                if (RecordDate.ToString("ddMMyy") == sixDays.ToString("ddMMyy"))
                                                {
                                                    daySix = MachineTotalVendMeter;
                                                }
                                                if (RecordDate.ToString("ddMMyy") == sevenDays.ToString("ddMMyy"))
                                                {
                                                    daySeven = MachineTotalVendMeter;
                                                }
                                                if (RecordDate.ToString("ddMMyy") == eightDays.ToString("ddMMyy"))
                                                {
                                                    dayEight = MachineTotalVendMeter;
                                                }
                                            }
                                        }
                                    //}
                                    //catch (Exception recordFailure)
                                    //{

                                        //Error.Add(recordFailure.ToString() + "_" + currentMachineNumber + " - " + currentDataFile + " - " + currentRecordDate);

                                    //}
                                }

                                //Timespan between oldest and newest record
                                TimeSpan DisAge = newest - oldest;

                                //Calculate daily average sales from data
                                int DailyAverage = (NewestReading - OldestReading) / DisAge.Days;

                                //Calculate weekly average sales
                                WeekAverage = (dayOne - daySeven) / 7;

                                //Calculate week average if 1 or more days are missing a reading
                                if (daySeven <= 0)
                                {
                                    WeekAverage = (dayOne - daySix) / 6;
                                }
                                if (daySix <= 0)
                                {
                                    WeekAverage = (dayOne - dayFive) / 5;
                                }
                                if (dayFive <= 0)
                                {
                                    WeekAverage = (dayOne - dayFour) / 4;
                                }
                                if (dayFour <= 0)
                                {
                                    WeekAverage = (dayOne - dayThree) / 3;
                                }

                                //Subtract day value from previous day value to get sales data from that day because the readings are accumalitive
                                dayOne = dayOne - dayTwo;
                                dayTwo = dayTwo - dayThree;
                                dayThree = dayThree - dayFour;
                                dayFour = dayFour - dayFive;
                                dayFive = dayFive - daySix;
                                daySix = daySix - daySeven;
                                daySeven = daySeven - dayEight;

                                //Calculate days to next scheduled visit
                                TimeSpan NextVisitDays = NextScheduledVisit - DateTime.Today;

                                //Calculate predicted stock count on scheduled visit day based on average sales
                                string NextVisitStock = (((DailyAverage * NextVisitDays.TotalDays) + Int32.Parse(StockSold.ToString())) / Int32.Parse(Capacity.ToString()) * 100).ToString("0");

                                //Decimal variables for calculations
                                decimal PercentSold1 = IntPercent;
                                decimal PercentSold2 = IntPercent;

                                //Int variables for calculations
                                int CountY = 0;
                                int CountX = 0;
                                int percentZ = 0;

                                //Calculate % from averages
                                if (StockSold < Capacity && DailyAverage > 0)
                                {
                                    while (PercentSold1 < 30)
                                    {
                                        PercentSold1 = PercentSold1 + ((decimal.Parse(DailyAverage.ToString()) / decimal.Parse(Capacity.ToString())) * 100);
                                        CountY++;
                                    }
                                }

                                //Days to stock sold @40%
                                if (StockSold < Capacity && DailyAverage > 0)
                                {
                                    while (PercentSold2 < 40)
                                    {
                                        PercentSold2 = PercentSold2 + ((decimal.Parse(DailyAverage.ToString()) / decimal.Parse(Capacity.ToString())) * 100); ;
                                        CountX++;
                                    }
                                }

                                //Predict stock level at next visit date
                                if (StockSold < Capacity && DailyAverage > 0)
                                {
                                    int StockSoldByNextVisit = (NextScheduledVisit - DateTime.Now).Days * DailyAverage;
                                    int StockSoldAtNextVisit = Int32.Parse(StockSold.ToString()) + StockSoldByNextVisit;
                                    percentZ = (int)(StockSoldAtNextVisit * Capacity) / 100;
                                    //If percent is more than 100, return it to 100
                                    if (percentZ > 100)
                                    {
                                        percentZ = 100;
                                    }
                                }

                                //If avg sales is unavailable, show % as 0 instead of 100
                                if (DailyAverage == 0)
                                {
                                    percentZ = 0;
                                }

                                //If % is more than 100, return it to 100
                                if (percentZ > 100)
                                {
                                    percentZ = 100;
                                }

                                //Check if Next visit stock prediction is more than 0, less than capacity and daily average is within threshold
                                if (Int32.Parse(NextVisitStock) > 0 && Int32.Parse(NextVisitStock) < Capacity && DailyAverage > 0 && DailyAverage < 400)
                                {
                                    //Assign variables to column rows
                                    row["Next Visit Stock Prediction"] = NextVisitStock + " %";
                                    row["Machine @30% Stock Sold In"] = CountY + " Days";
                                    row["Machine @40% Stock Sold In"] = CountX + " Days";
                                    row["Optimal Fill Date"] = DateTime.Today.AddDays(CountX).ToString("dd MMM yy");
                                    row["Optimal Fill Date In:"] = CountX.ToString() + " Days";
                                    row["Optimal Fill Day"] = DateTime.Today.AddDays(CountX).DayOfWeek;
                                }
                                else
                                {
                                    //If data does not fit within threshold then set the row value to indicate the missing or incorrect data
                                    row["Machine @30% Stock Sold In"] = "Missing Data";
                                    row["Machine @40% Stock Sold In"] = "Missing Data";
                                    row["Next Visit Stock Prediction"] = "Missing Data";
                                    row["Optimal Fill Date"] = "Missing Data";
                                    row["Optimal Fill Date In:"] = "Missing Data";
                                    row["Optimal Fill Day"] = "Missing Data";
                                }

                                //Check if daily average is within threshold and assign the column row if it is
                                if (DailyAverage > 0 && DailyAverage < 400)
                                {
                                    row["Average Sales Per Week"] = DailyAverage.ToString();
                                }
                                else
                                {
                                    //Else assign the column row to indicate missing or incorrect data
                                    row["Average Sales Per Week"] = "Missing Data";
                                }

                                //Assign varaibles to the columns of the previous 7 day's sales
                                row[yesterday.DayOfWeek.ToString() + "|" + yesterday.ToString("dd-MMM-yy")] = dayTwo;
                                row[twoDays.DayOfWeek.ToString() + "|" + twoDays.ToString("dd-MMM-yy")] = dayThree;
                                row[threeDays.DayOfWeek.ToString() + "|" + threeDays.ToString("dd-MMM-yy")] = dayFour;
                                row[fourDays.DayOfWeek.ToString() + "|" + fourDays.ToString("dd-MMM-yy")] = dayFive;
                                row[fiveDays.DayOfWeek.ToString() + "|" + fiveDays.ToString("dd-MMM-yy")] = daySix;
                                row[sixDays.DayOfWeek.ToString() + "|" + sixDays.ToString("dd-MMM-yy")] = daySeven;
                                row[sevenDays.DayOfWeek.ToString() + "|" + sevenDays.ToString("dd-MMM-yy")] = dayEight;
                            }
                        }
                    //}
                    //catch (Exception fileFailure)
                    //{
                        //Catch any errors and add to error log list
                        //Error.Add(fileFailure.ToString() + "_" + currentMachineNumber + " - " + currentDataFile + " - " + currentRecordDate);
                    //}
                }
                catch (Exception batchFailure)
                {
                    Error.Add(batchFailure.ToString() + "_" + currentMachineNumber + " - " + currentDataFile + " - " + currentRecordDate);
                }

                //Add rows to data table
                Data.Rows.Add(row);
            }

            //Add data table to dataset for ecporting to Excel sheet
            dataSet.Tables.Add(Data);

            //Export DataSet to .CSV

            //Get date/time of previous Excel sheet
            DateTime Yesterday = File.GetLastWriteTime(Output + "Ispy-Report.xls");
            
            //Move previous Excel sheet and add date/time of report to the file name
            File.Move(Output + "Ispy-Report.xls", OldReports + "Old_Report-" + Yesterday.ToString("dd_MM_yy~HH;mm;ss") + ".xls");

            //Export dataset to Excel sheet using DocumentFormat NuGet package excel template export method
            ExportDataSet(dataSet, Output + "Ispy-Report.xls");

            //If any errors occur write log file to log file folder
            if (Error.Count > 0)
            {
                File.WriteAllLines(ErrorLog + @"File_Operation_Log\" + "File_Log_" + DateTime.Now.ToString("dd-MMM-yyyy_HH-mm-ss") + ".txt", Error);
            }

            //Output to console that the report has completed and include a count of the errors that occured
            Console.WriteLine("Report completed with " + Error.Count + " errors");
        }
    }
}
