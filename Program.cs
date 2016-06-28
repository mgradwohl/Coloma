using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Win32;

namespace Coloma
{
    class Program
    {
        // Just load deviceId once, it's independant of other operations.
        private static readonly string DeviceId = GetDeviceId();

        static void Main(string[] args)
        {
            // bail out if the user isn't running Windows 10
            if (Environment.OSVersion.Version.Major < 10)
            {
                Console.WriteLine();
                Console.WriteLine("Coloma requires Windows 10");
                Console.WriteLine();
                Console.WriteLine("Hit any key to exit");
                Console.ReadKey(true);

                return;
            }

            // Inform the user we're running
            Console.WriteLine();
            Console.WriteLine("Coloma is gathering log entries for your machine.");
            if (DeviceId.Contains("UNKNOWN"))
            {
                Console.WriteLine("WARNING: Coloma could not automatically detect your DeviceID.");
            }
            Console.WriteLine();

            // create the file on the network share, unless it's unavailable, then use the desktop
            string filename = Environment.MachineName + "_" + Environment.UserName + "_" + Environment.TickCount.ToString() + ".tsv";
            string filepath = @"\\iefs\users\mattgr\Coloma" + "\\Coloma" + "_" + filename;
            StreamWriter sw;
            try
            {
                sw = new StreamWriter(filepath, false, System.Text.Encoding.UTF8);
            }
            catch (Exception)
            {
                filepath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Coloma" + "_" + filename;
                sw = new StreamWriter(filepath, false, System.Text.Encoding.UTF8);
            }

            // just get logs since last time OR since 4/1/2016
            DateTime dt = new DateTime(2016, 4, 1, 0, 0, 0, 0, DateTimeKind.Local);
            GetLastColomaDate(ref dt);

            // Tell the user what we're doing
            Console.WriteLine("Any error, warning, or KB install written after " + dt.ToShortDateString());
            Console.WriteLine("From the following logs: System, Security, Hardware Events, Setup, and Application");
            Console.WriteLine("Data will be saved to " + filepath);
            Console.WriteLine();

            Console.Write("KB Articles for revision install history... ");
            // TH2 revisions
            List<KBRevision> kbrlist = new List<KBRevision>
            {
                new KBRevision(10586, 420, "KB3163018"),
                new KBRevision(10586, 318, "KB3156421"),
                new KBRevision(10586, 218, "KB3147458"),
                new KBRevision(10586, 164, "KB3140768"),
                new KBRevision(10586, 122, "KB3140743"),
                new KBRevision(10586, 104, "KB3135173"),
                new KBRevision(10586, 71, "KB3124262"),
                new KBRevision(10586, 63, "KB3124263"),
                new KBRevision(10586, 36, "KB3124200"),
                new KBRevision(10586, 29, "KB3116900"),
                new KBRevision(10586, 17, "KB3116908"),
                new KBRevision(10586, 14, "KB3120677"),
                new KBRevision(10586, 11, "KB3118754")
            };
            Console.WriteLine("done");

            List<ColomaEvent> eventlist = new List<ColomaEvent>();

            // Setup logs are different than other logs
            Console.Write("Setup... ");
            bool setuplog = AddSetupLogToList(kbrlist, eventlist, dt);
            Console.WriteLine("done");

            // Go through the 'standard' logs
            string[] Logs = { "System", "HardwareEvents", "Application", "Security" };
            foreach (string log in Logs)
            {
                EventLog eventlog = new EventLog(log, ".");
                Console.Write(log + "... ");
                AddStandardLogToList(eventlog, eventlist, dt);
                eventlog.Close();
                Console.WriteLine("done");
            }

            // gets all the events in order, and uses info from the setup log to ensure the correct revision
            // if the setuplog had no entries, then use the current build and revision
            Console.Write("Sort and fixup... ");
            SortandFix(eventlist, setuplog);
            Console.WriteLine("done");

            Console.Write("Writing file");
            int i = 0;
            sw.WriteLine(ColomaEvent.Header());
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
            Console.ReadKey(true);
        }

        private static void SortandFix(List<ColomaEvent> eventlist, bool setuplog)
        {
            uint r = 11;
            if (!setuplog)
            {
                WindowsVersion.WindowsVersionInfo wvi = new WindowsVersion.WindowsVersionInfo();
                WindowsVersion.GetWindowsBuildandRevision(wvi);
                r = wvi.revision;
            }

            // sort the list by date
            eventlist.Sort();

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

        private static bool AddSetupLogToList(List<KBRevision> kbrlist, List<ColomaEvent> list, DateTime dt)
        {
            // this retrieves the build.revision and branch for the current client
            WindowsVersion.WindowsVersionInfo wvi = new WindowsVersion.WindowsVersionInfo();
            WindowsVersion.GetWindowsBuildandRevision(wvi);
            uint revision = wvi.revision;
            bool setuplog = false;

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
                        // this is a KB installed message, figure out which KB it is and update the revision of that entry
                        string kb = "KB";
                        int i = msg.IndexOf(kb);
                        setuplog = true;

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
                    list.Add(new ColomaEvent(wvi.branch, wvi.build, revision, entry.MachineName, DeviceId,
                                             Environment.UserName, "Setup", entry.LevelDisplayName, entry.Id,
                                             entry.TimeCreated.GetValueOrDefault(), entry.ProviderName, msg));
                }
            }
            return setuplog;
        }

        private static void AddStandardLogToList(EventLog log, List<ColomaEvent> list, DateTime dt)
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
                        list.Add(new ColomaEvent(wvi.branch, wvi.build, 0, entry.MachineName, DeviceId,
                                 Environment.UserName, log.Log/*.LogDisplayName*/, entry.EntryType.ToString(),
                                 entry.InstanceId, entry.TimeGenerated, entry.Source, msg));
                    }
                }
            }
        }

        private static string CleanUpMessage(string message)
        {
            string msg = message.Replace("\t", " ");
            msg = msg.Replace("\r\n", "<br>");
            msg = msg.Replace("\n", "<br>");
            msg = msg.Replace("\r", "<br>");
            msg = msg.Replace("<br><br>", "<br>");

            return msg;
        }

        private static void GetLastColomaDate(ref DateTime dtLastDate)
        {
            const string keyName = "SOFTWARE\\Coloma";
            const string valueName = "LastLogDate";

            RegistryKey rk = Registry.CurrentUser.CreateSubKey(keyName, true);

            try
            {
                long dtl = (long)rk.GetValue(valueName, dtLastDate.ToBinary());
                dtLastDate = DateTime.FromBinary(dtl);
            }
            catch (ArgumentException)
            {
                // ArgumentException is thrown if the key does not exist. In
                // this case, there is no reason to display a message.
            }

            rk.SetValue(valueName, DateTime.Now.ToBinary(), RegistryValueKind.QWord);

            if (rk != null)
            {
                rk.Close();
            }
        }

        // Attempts to retrieve the deviceID from the registry
        private static string GetDeviceId()
        {
            const string keyName =
                "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Diagnostics\\DiagTrack\\SettingsRequests\\CFC.FLIGHTS";

            const string valueName = @"ETagQueryParameters";

            RegistryKey rk = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
            rk = rk.OpenSubKey(keyName);

            // We need to see if the key exists or contains what we need.
            if (rk != null)
            {
                string keyData = (string)rk.GetValue(valueName);

                if (keyData.Contains("deviceId"))
                {
                    // I'm not sure of deviceID's form, so to be safer instead of using substring, I will use regex.
                    // deviceId=s:57D3C3B8-F144-4E2D-8443-3BF3D95CB5DA&
                    var regex = new Regex(@"deviceId=(.*?)&");

                    // See if we matched anything
                    if (regex.IsMatch(keyData))
                    {
                        // Send back the deviceId
                        return regex.Match(keyData).Groups[1].Value;
                    }
                }
            }

            // If we fell out of the any of the above, deviceID is unknown
            return "s:UNKNOWN";
        }
    }
}