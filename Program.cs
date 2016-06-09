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
            // bail out if the user isn't running Windows 10
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

            Console.Write("KB Articles... ");
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
            Console.WriteLine("done");

            List<ColomaEvent> eventlist = new List<ColomaEvent>();

            Console.Write("Setup... ");
            // this will also fill in the list of revisions so we know when a build was updated
            AddSetupLogToList(kbrlist, eventlist, dt);
            Console.WriteLine("done");

            // one log at a time
            string[] Logs = { "System", "HardwareEvents", "Application", "Security" };
            foreach (string log in Logs)
            {
                EventLog eventlog = new EventLog(log, ".");
                Console.Write(log + "... ");
                AddStandardLogToList(eventlog, eventlist, dt);
                eventlog.Close();
                Console.WriteLine("done");
            }

            Console.Write("Sort and fixup... ");
            SortandFix(eventlist);
            Console.WriteLine("done");

            Console.Write("Writing file");
            int i = 0;
            foreach (ColomaEvent evt in eventlist)
            {
                i++;
                sw.WriteLine(evt.ToString());
                if (i % 10 == 0)
                {
                    Console.Write(".");
                }
            }
            Console.WriteLine(" done");
            sw.Close();

            Console.WriteLine();
            Console.WriteLine("Done, thank you. Hit any key to exit");
            Console.ReadLine();
        }

        private static void SortandFix(List<ColomaEvent> eventlist)
        {
            // sort the list by date
            eventlist.Sort();
            uint r = 11;
            foreach (ColomaEvent evt in eventlist)
            {
                if (evt.Logname == "Setup")
                {
                    r = evt.Revision;
                }
                else
                {
                    evt.Revision = r;
                }
            }
        }

        static void AddSetupLogToList(List<KBRevision> kbrlist, List<ColomaEvent> list, DateTime dt)
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
                        string kb = "KB";
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
                            }
                        }
                    }
                    list.Add(new ColomaEvent(wvi.branch, wvi.build, revision, entry.MachineName, Environment.UserName, "Setup", entry.LevelDisplayName, entry.TimeCreated.GetValueOrDefault(), entry.ProviderName, msg));
                }
            }
        }

        static void AddStandardLogToList(EventLog log, List<ColomaEvent> list, DateTime dt)
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
                        list.Add(new ColomaEvent(wvi.branch, wvi.build, 0, entry.MachineName, Environment.UserName, log.LogDisplayName, entry.EntryType.ToString(), entry.TimeGenerated, entry.Source, msg));
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