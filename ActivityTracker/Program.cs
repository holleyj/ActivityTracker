using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using NLog;

namespace ActivityTracker
{
    class Program
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private static Dictionary<string, DateTime> activeActivities = new Dictionary<string, DateTime>();
        private static Dictionary<string, DateTime> pausedActivities = new Dictionary<string, DateTime>();
        private static List<Activity> activityLog = new List<Activity>();
        private const string LogFilePath = @"C:\log\timeReport.csv";

        static void Main(string[] args)
        {
            Logger.Info("Activity Tracker Started");
            string prompt = string.Empty;

            while (prompt != "exit")
            {
                Console.WriteLine("Enter command (start/stop/pause/restart/add/list/open/exit):");
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
                    case "pause":
                        PauseActivity(arg);
                        break;
                    case "restart":
                        RestartActivity(arg);
                        break;
                    case "add":
                        AddActivity();
                        break;
                    case "list":
                        ListActivities();
                        break;
                    case "open":
                        OpenLogFile();
                        break;
                    case "save":
                        GenerateReport();
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


            ListActivities();
        }

        private static void PauseActivity(string arg)
        {
            string input;
            if (arg == string.Empty)
            {
                Console.WriteLine("Enter the number of the activity to pause:");
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
                DateTime pauseTime = DateTime.Now;
                double hoursSpent = Math.Round((pauseTime - startTime).TotalHours, 2);

                pausedActivities[activityName] = pauseTime;

                activeActivities.Remove(activityName);
                Console.WriteLine($"Paused activity: {activityName}. Time spent: {hoursSpent} hours so far");
                Logger.Info($"Paused activity: {activityName}. Time spent: {hoursSpent} hours so far");
            }
            else
            {
                Console.WriteLine("Invalid activity number.");
            }

            ListActivities();
        }

        private static void RestartActivity(string arg)
        {
            string input;
            if (arg == string.Empty)
            {
                Console.WriteLine("Enter the number of the activity to restart:");
                input = Console.ReadLine();
            }
            else
            {
                input = arg;
            }
            if (int.TryParse(input, out int activityNumber) && activityNumber > 0 && activityNumber <= pausedActivities.Count)
            {
                var activityName = new List<string>(pausedActivities.Keys)[activityNumber - 1];
                DateTime pauseTime = pausedActivities[activityName];

                activeActivities[activityName] = pauseTime;

                pausedActivities.Remove(activityName);
                Console.WriteLine($"Restarted activity: {activityName}");
                Logger.Info($"Restarted activity: {activityName}");
            }
            else
            {
                Console.WriteLine("Invalid activity number.");
            }

            ListActivities();
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
                Console.WriteLine("\nNo active activities.");
            }
            else
            {
                Console.WriteLine("\nActive activities:");
            }

            int index = 1;
            foreach (var activity in activeActivities)
            {
                Console.WriteLine($"{index}. {activity.Key} (Started at: {activity.Value})");
                index++;
            }

            Console.WriteLine("");

            if (pausedActivities.Count > 0)
            {
                Console.WriteLine("Paused activities:");
                index = 1;
                foreach (var activity in pausedActivities)
                {
                    Console.WriteLine($"{index}. {activity.Key} (Paused at: {activity.Value})");
                    index++;
                }
                Console.WriteLine("");
            }
        }

        private static void OpenLogFile()
        {
            try
            {
                Process.Start(@"C:\Program Files\Notepad++\notepad++.exe", LogFilePath);
                Logger.Info("Opened log file with Notepad++");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error opening log file: {ex.Message}");
                Logger.Error($"Error opening log file: {ex.Message}");
            }
        }

        private static void GenerateReport()
        {
            int index = activeActivities.Count;
            var loopActivities = new Dictionary<string, DateTime>(activeActivities);
            foreach (var activity in loopActivities)
            {
                StopActivity(index.ToString());
                index--;
            }

            try
            {
                List<string> lines = new();
                foreach (var activity in activityLog)
                {
                    string row = $"{activity.Name},{activity.StartTime},{activity.EndTime},{activity.HoursSpent}";
                    lines.Add(row);
                }
                File.AppendAllLines(LogFilePath, lines);

                Console.WriteLine($"Report generated: {LogFilePath}");
                Logger.Info($"Report generated: {LogFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error generating report: {ex.Message}");
                Logger.Error($"Error generating report: {ex.Message}");
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
