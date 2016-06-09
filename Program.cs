using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.IO;

namespace Coloma
{
    class Program
    {
        static void Main(string[] args)
        {
            if (Environment.OSVersion.Version.Major < 10)
            {
                Console.WriteLine();
                Console.WriteLine("Coloma requires Windows 10");
                Console.WriteLine();
                Console.WriteLine("Hit any key to exit");
                Console.ReadLine();

                return;
            }

            // Inform the user we're running
            Console.WriteLine();
            Console.WriteLine("Coloma is gathering log entries");
            Console.WriteLine();

            List<KBRevision> kbrlist = new List<KBRevision>();
            kbrlist.Add(new KBRevision(10586, 318, "KB3156421"));
            kbrlist.Add(new KBRevision(10586, 218, "KB3147458"));
            kbrlist.Add(new KBRevision(10586, 164, "KB3140768"));
            kbrlist.Add(new KBRevision(10586, 122, "KB3140743"));
            kbrlist.Add(new KBRevision(10586, 104, "KB3135173"));
            kbrlist.Add(new KBRevision(10586, 71, "KB3124262"));
            kbrlist.Add(new KBRevision(10586, 63, "KB3124263"));
            kbrlist.Add(new KBRevision(10586, 36, "KB3124200"));
            kbrlist.Add(new KBRevision(10586, 29, "KB3116900"));
            kbrlist.Add(new KBRevision(10586, 17, "KB3116908"));
            kbrlist.Add(new KBRevision(10586, 14, "KB3120677"));
            kbrlist.Add(new KBRevision(10586, 11, "KB3118754"));

            // create the file on the network share, unless it's unavailable, then use the desktop
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


            // Tell the user what we're doing
            Console.WriteLine("Any error, warning, or KB install written after " + dt.ToShortDateString());
            Console.WriteLine("From the following logs: system, security, hardwareevents, setup, and application");
            Console.WriteLine("Data will be saved to " + filename);
            Console.WriteLine();

            Console.Write("Setup... ");
            // this will also fill in the list of revisions so we know when a build was updated
            WriteSetupLogToStream(kbrlist, sw, dt);
            Console.WriteLine("done");

            // one log at a time
            string[] Logs = { "System", "HardwareEvents", "Application", "Security" };
            foreach (string log in Logs)
            {
                EventLog eventlog = new EventLog(log, ".");
                Console.Write(log + "... ");
                WriteLogToStream(eventlog, sw, dt);
                eventlog.Close();
                Console.WriteLine("done");
            }

            sw.Close();

            Console.WriteLine();
            Console.WriteLine("Done, thank you. Hit any key to exit");
            Console.ReadLine();
        }

        static void WriteSetupLogToStream(List<KBRevision> kbrlist, StreamWriter sw, DateTime dt)
        {
            // this retrieves the build.revision and branch for the current client
            WindowsVersion.WindowsVersionInfo wvi = new WindowsVersion.WindowsVersionInfo();
            WindowsVersion.GetWindowsBuildandRevision(wvi);
            uint revision = wvi.revision;

            EventLogQuery query = new EventLogQuery("Setup", PathType.LogName);
            query.ReverseDirection = false;
            EventLogReader reader = new EventLogReader(query);

            EventRecord entry;
            while ((entry = reader.ReadEvent()) != null)
            {
                if ((entry.Level == (byte)StandardEventLevel.Critical) ||
                    (entry.Level == (byte)StandardEventLevel.Error) ||
                    (entry.Level == (byte)StandardEventLevel.Warning) ||
                    (entry.Id == 2))
                {
                    string msg = CleanUpMessage(entry.FormatDescription());

                    if (entry.Id == 2)
                    {
                        // this is a KB installed message, figure out which KB it is and update the revision
                        string kb = "KB31";
                        int i = msg.IndexOf(kb);

                        if (-1 != i)
                        {
                            // we found the kb article
                            kb = msg.Substring(i, 9);

                            foreach (KBRevision rev in kbrlist)
                            {
                                if (rev.Kb == kb)
                                {
                                    revision = rev.Revision;
                                }
                                else
                                {
                                    // unknown
                                    revision = 0xffff;
                                }
                            }
                        }
                    }
                    sw.WriteLine("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}", wvi.branch, wvi.build.ToString(), revision.ToString(), entry.MachineName, Environment.UserName, "Setup", entry.LevelDisplayName, entry.TimeCreated.ToString(), entry.ProviderName, msg);
                }
            }
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
                        sw.WriteLine("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}", wvi.branch, wvi.build.ToString(), wvi.revision.ToString(), entry.MachineName, Environment.UserName, log.LogDisplayName, entry.EntryType.ToString(), entry.TimeGenerated.ToString(), entry.Source, msg);
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
