using NLog;
using System.Data;

namespace ActivityTracker
{
    class Program
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private static Dictionary<string, DateTime> activeActivities = new Dictionary<string, DateTime>();
        private static List<Activity> activityLog = new List<Activity>();
        private const string LogFilePath = @"C:\log\timeReport.csv";

        static void Main(string[] args)
        {
            Logger.Info("Activity Tracker Started");
            string prompt = string.Empty;

            while (prompt != "exit")
            {
                Console.WriteLine("Enter command (start/stop/add/list/exit):");
                prompt = Console.ReadLine();

                var x = prompt.Split(" ");
                string command = x[0];
                string arg = string.Empty;

                if (x.Length == 2)
                {
                    arg = x[1];
                }


                switch (command.ToLower())
                {
                    case "start":
                        StartActivity(arg);
                        break;
                    case "stop":
                        StopActivity(arg);
                        break;
                    case "add":
                        AddActivity();
                        break;
                    case "list":
                        ListActivities();
                        break;
                    case "exit":
                        GenerateReport();
                        Logger.Info("Activity Tracker Stopped");
                        LogManager.Shutdown();
                        break;
                    default:
                        Console.WriteLine("Invalid command.");
                        break;
                }
            }
        }

        private static void StartActivity(string arg)
        {
            string activityName;
            if (arg == string.Empty)
            {
                Console.WriteLine("Enter activity name:");
                activityName = Console.ReadLine();
            }
            else
            {
                activityName = arg;
            }
            if (activeActivities.ContainsKey(activityName))
            {
                Console.WriteLine("This activity is already started.");
            }
            else
            {
                activeActivities[activityName] = DateTime.Now;
                Console.WriteLine($"Started activity: {activityName}");
                Logger.Info($"Started activity: {activityName}");
            }

            ListActivities();
        }

        private static void StopActivity(string arg)
        {
            ListActivities();
            string input;
            if (arg == string.Empty)
            {
                Console.WriteLine("Enter the number of the activity to stop:");
                input = Console.ReadLine();
            }
            else
            {
                input = arg;
            }
            if (int.TryParse(input, out int activityNumber) && activityNumber > 0 && activityNumber <= activeActivities.Count)
            {
                var activityName = new List<string>(activeActivities.Keys)[activityNumber - 1];
                DateTime startTime = activeActivities[activityName];
                DateTime endTime = DateTime.Now;
                double hoursSpent = Math.Round((endTime - startTime).TotalHours, 2);
                

                activityLog.Add(new Activity
                {
                    Name = activityName,
                    StartTime = startTime,
                    EndTime = endTime,
                    HoursSpent = hoursSpent
                });

                activeActivities.Remove(activityName);
                Console.WriteLine($"Stopped activity: {activityName}. Time spent: {hoursSpent} hours");
                Logger.Info($"Stopped activity: {activityName}. Time spent: {hoursSpent} hours");
            }
            else
            {
                Console.WriteLine("Invalid activity number.");
            }
        }

        private static void AddActivity()
        {
            Console.WriteLine("Enter activity name:");
            string activityName = Console.ReadLine();
            Console.WriteLine("Enter time spent (hours):");
            if (double.TryParse(Console.ReadLine(), out double hoursSpent))
            {
                hoursSpent = Math.Round(hoursSpent, 2);
                activityLog.Add(new Activity
                {
                    Name = activityName,
                    StartTime = DateTime.Now,
                    EndTime = DateTime.Now,
                    HoursSpent = hoursSpent
                });
                Console.WriteLine($"Added activity: {activityName}. Time spent: {hoursSpent} hours");
                Logger.Info($"Added activity: {activityName}. Time spent: {hoursSpent} hours");
            }
            else
            {
                Console.WriteLine("Invalid time format.");
            }
        }

        private static void ListActivities()
        {
            if (activeActivities.Count == 0)
            {
                Console.WriteLine("No active activities.");
                return;
            }

            Console.WriteLine("\nActive activities:");
            int index = 1;
            foreach (var activity in activeActivities)
            {
                Console.WriteLine($"{index}. {activity.Key} (Started at: {activity.Value})");
                index++;
            }
            Console.WriteLine("");
        }

        private static void GenerateReport()
        {
            int index = activeActivities.Count;
            var loopActivities = activeActivities;
            foreach (var activity in loopActivities)
            {
                Console.WriteLine(index);
                StopActivity(index.ToString());
                index--;
            }
            //Application excelApp = new Application();
            //Workbook workbook = null;
            //Worksheet worksheet = null;

            try
            {
                //workbook = excelApp.Workbooks.Add();
                //worksheet = (Worksheet)workbook.Worksheets[1];
                //worksheet.Name = "Activity Log";

                //((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[1, 1]).Value = "Activity Name";
                //((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[1, 2]).Value = "Start Time";
                //((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[1, 3]).Value = "End Time";
                //((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[1, 4]).Value = "Hours Spent";

                List<string> lines = new();
                foreach (var activity in activityLog)
                {

                    string row = $"{activity.Name},{activity.StartTime.ToString()},{activity.EndTime.ToString()},{activity.HoursSpent}";
                    lines.Add(row);
                    //((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[row, 1]).Value = activity.Name;
                    //((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[row, 2]).Value = activity.StartTime.ToString();
                    //((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[row, 3]).Value = activity.EndTime.ToString();
                    //((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[row, 4]).Value = activity.HoursSpent;
                    //row++;
                }
                File.AppendAllLines(LogFilePath, lines);

                //workbook.SaveAs(LogFilePath);
                Console.WriteLine($"Report generated: {LogFilePath}");
                Logger.Info($"Report generated: {LogFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error generating report: {ex.Message}");
                Logger.Error($"Error generating report: {ex.Message}");
            }
            finally
            {
                //if (workbook != null)
                //{
                //    workbook.Close();
                //    Marshal.ReleaseComObject(workbook);
                //}

                //if (excelApp != null)
                //{
                //    excelApp.Quit();
                //    Marshal.ReleaseComObject(excelApp);
                //}
            }
        }
    }

    public class Activity
    {
        public string Name { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public double HoursSpent { get; set; }
    }
}
