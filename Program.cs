using System;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.IO;

namespace Coloma
{
    class Program
    {
        static void Main(string[] args)
        {
            // create the file
            string filename = @"\\iefs\users\mattgr\Coloma" + "\\Coloma" + "_" + Environment.MachineName + "_" + Environment.UserName + "_" + Environment.TickCount.ToString() + ".csv";
            StreamWriter sw;
            try
            {
                sw = new StreamWriter(filename, false, System.Text.Encoding.UTF8);
            }
            catch (Exception)
            {
                filename = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Coloma" + "_" + Environment.MachineName + "_" + Environment.UserName + "_" + Environment.TickCount.ToString() + ".csv";
                sw = new StreamWriter(filename, false, System.Text.Encoding.UTF8);
            }

            // just get logs for 3/1/2016 and after
            DateTime dt = new DateTime(2016, 3, 1, 0, 0, 0, 0, DateTimeKind.Local);

            Console.WriteLine();
            Console.WriteLine("Coloma is gathering log entries");
            Console.WriteLine();
            Console.WriteLine("Any error, warning, or KB install written after " + dt.ToShortDateString());
            Console.WriteLine("From the following logs: system, security, hardwareevents, setup, and application");
            Console.WriteLine("Data will be saved to " + filename);
            Console.WriteLine();

            string[] Logs = { "System", "HardwareEvents", "Application", "Security" };

            // one log at a time
            foreach (string log in Logs)
            {
                EventLog eventlog = new EventLog(log, ".");
                Console.Write(log + "... ");
                WriteLogToStream(eventlog, sw, dt);
                eventlog.Close();
                Console.WriteLine("done");
            }

            Console.Write("Setup... ");
            WriteSetupLogToStream(sw, dt);
            Console.WriteLine("done");

            sw.Close();

            Console.WriteLine();
            Console.WriteLine("Done, thank you. Hit any key to exit");
            Console.ReadLine();
        }

        static void WriteLogToStream(EventLog log, StreamWriter sw, DateTime dt)
        {
            WindowsVersion.WindowsVersionInfo wvi = new WindowsVersion.WindowsVersionInfo();
            WindowsVersion.GetWindowsBuildandRevision(wvi);
            
            foreach (EventLogEntry entry in log.Entries)
            {
                if (entry.TimeGenerated > dt)
                {
                    if ((entry.EntryType == EventLogEntryType.Error) ||
                        (entry.EntryType == EventLogEntryType.Warning))
                    {
                        string msg = CleanUpMessage(entry.Message);
                        sw.WriteLine("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}", wvi.branch, wvi.build.ToString(), wvi.revision.ToString(), Environment.MachineName, Environment.UserName, log.LogDisplayName, entry.EntryType.ToString(), entry.TimeGenerated.ToString(), entry.Source, msg);
                    }
                }
            }
        }

        static void WriteSetupLogToStream(StreamWriter sw, DateTime dt)
        {
            WindowsVersion.WindowsVersionInfo wvi = new WindowsVersion.WindowsVersionInfo();
            WindowsVersion.GetWindowsBuildandRevision(wvi);

            EventLogQuery query = new EventLogQuery("Setup", PathType.LogName);
            query.ReverseDirection = false;
            EventLogReader reader = new EventLogReader(query);

            EventRecord entry;
            while ((entry = reader.ReadEvent()) != null)
            {
                if (entry.TimeCreated > dt)
                {
                    if ((entry.Level == (byte)StandardEventLevel.Critical) ||
                        (entry.Level == (byte)StandardEventLevel.Error) ||
                        (entry.Level == (byte)StandardEventLevel.Warning) ||
                        (entry.Id == 2))
                    {
                        string msg = CleanUpMessage(entry.FormatDescription());
                        sw.WriteLine("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}", wvi.branch, wvi.build.ToString(), wvi.revision.ToString(), Environment.MachineName, Environment.UserName, "Setup", entry.LevelDisplayName, entry.TimeCreated.ToString(), entry.ProviderName, msg);
                    }
                }
            }
        }

        static string CleanUpMessage(string Message)
        {
            string msg = Message.Replace("\t", " ");
            msg = msg.Replace("\r\n", "<br>");
            msg = msg.Replace("\n", "<br>");
            msg = msg.Replace("<br><br>", "<br>");

            return msg;
        }
    }
}
