using System;
using System.Reflection;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Win32;
using System.Diagnostics;

namespace Coloma
{
    public class Coloma
    {
        // Just load deviceId once, it's independant of other operations.
        private static readonly string DeviceId = GetDeviceId();

        private static readonly string descNotFoundMsgTemplate = "The description for Event ID '{0}' in Source '{1}' cannot be found.  The local computer may not have the necessary registry information or message DLL files to display the message, or you may not have permission to access them.  The following information is part of the event:{2}";
        private static readonly string structuredQueryTemplate = "<QueryList> <Query Id=\"0\"> <Select Path=\"Setup\">*[System[(Level=1 or Level=2 or Level=3) or (EventID=2)]]</Select> <Select Path=\"Application\">*[System[(Level=1 or Level=2 or Level=3) and TimeCreated[@SystemTime&gt;='{0:O}']]]</Select> <Select Path=\"System\">*[System[(Level=1 or Level=2 or Level=3) and TimeCreated[@SystemTime&gt;='{0:O}']]]</Select> <Select Path=\"Security\">*[System[(Level=1 or Level=2 or Level=3) and TimeCreated[@SystemTime&gt;='{0:O}']]]</Select> <Select Path=\"HardwareEvents\">*[System[(Level=1 or Level=2 or Level=3) and TimeCreated[@SystemTime&gt;='{0:O}']]]</Select> </Query> </QueryList>";

        private static uint defaultRevision = 11;
        // TH2 revisions
        private static List<KBRevision> kbrlist = new List<KBRevision>
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

            // Parse args
            var onlyNewEvents = !args.Contains("-all");
            var saveLocal = args.Contains("-local");

            // Inform the user we're running
            Console.WriteLine();
            Console.WriteLine("Coloma is gathering log entries for your machine.");
            Console.WriteLine();

            // just get logs since last time OR since 4/1/2016
            DateTime dt = new DateTime(2016, 4, 1, 0, 0, 0, 0, DateTimeKind.Local);
            if (onlyNewEvents)
            {
                GetLastColomaDate(ref dt);
            }

            // Tell the user what we're doing
            Console.WriteLine("Any error, warning, or KB install written after " + dt.ToShortDateString());
            Console.WriteLine("From the following logs: System, Security, Hardware Events, Setup, and Application");

            if (DeviceId.Contains("UNKNOWN"))
            {
                Console.WriteLine("WARNING: Coloma could not automatically detect your DeviceID.");
            }

            Console.WriteLine();

            List<ColomaEvent> eventlist = new List<ColomaEvent>();

            Console.WriteLine("Collecting events...");

            Stopwatch timer = Stopwatch.StartNew();
            AddLogToList(eventlist, dt);
            timer.Stop();

            Console.WriteLine("done. {0} events in {1}ms", eventlist.Count, timer.ElapsedMilliseconds);

            if (eventlist.Count > 0)
            {

                // create the file on the network share, unless it's unavailable, then use the desktop
                string filename = Assembly.GetExecutingAssembly().GetName().Version.ToString() + "_" + Environment.MachineName + "_" + Environment.UserName + "_" + Environment.TickCount.ToString() + ".tsv";
                string filepath = @"\\iefs\users\mattgr\Coloma" + "\\Coloma" + "_" + filename;
                StreamWriter sw;
                try
                {
                    if (saveLocal) throw new Exception();
                    sw = new StreamWriter(filepath, false, System.Text.Encoding.UTF8);
                }
                catch (Exception)
                {
                    filepath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Coloma" + "_" + filename;
                    sw = new StreamWriter(filepath, false, System.Text.Encoding.UTF8);

                    Console.WriteLine(saveLocal
                        ? "Coloma will write the .tsv to your desktop."
                        : "Coloma could not access the network share and will write the .tsv to your desktop.");
                }

                Console.Write("Writing file to " + filepath);

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
            }

            Console.WriteLine();
            Console.WriteLine("Done, thank you. Hit any key to exit");
            Console.ReadKey(true);
        }

        private static void AddLogToList(List<ColomaEvent> list, DateTime dt)
        {
            // this retrieves the build.revision and branch for the current client
            WindowsVersion.WindowsVersionInfo wvi = new WindowsVersion.WindowsVersionInfo();
            WindowsVersion.GetWindowsBuildandRevision(wvi);
            uint revision = wvi.revision;

            int firstKbEvent = 0;

            Guid servicingProvider = new Guid("BD12F3B8-FC40-4A61-A307-B7A013A069C1");
            string structuredQuery = string.Format(System.Globalization.CultureInfo.InvariantCulture, Coloma.structuredQueryTemplate, dt.ToUniversalTime());

            EventLogQuery query = new EventLogQuery(null, PathType.LogName, structuredQuery);
            query.ReverseDirection = false;
            EventLogReader reader = new EventLogReader(query);
            // The Event Log can only return a maximum of 2 MB at a time, but it does not actually limit itself to this when collecting <BatchSize> events.
            // If the number of requested events exceeds 2 MB of event data an exception will be throw (System.Diagnostics.Eventing.Reader.EventLogException: The data area passed to a system call is too small).
            // Since an event is at most 64 KB, 30 events is a conservative limit to ensure the 2 MB limit is never crossed.
            // Setting a smaller batch size does have a small performance impact, but not enough to notice in this scenario.
            reader.BatchSize = 30;

            EventRecord entry;
            while ((entry = reader.ReadEvent()) != null)
            {
                var entryTimeCreated = entry.TimeCreated ?? new DateTime();
                var entryMachineName = entry.MachineName;
                var entryLogName = entry.LogName;
                var entryId = entry.Id;
                var entryProviderName = entry.ProviderName;

                string entryLevelDisplayName = "";
                try
                {
                    entryLevelDisplayName = entry.LevelDisplayName;
                }
                catch
                {
                    if (entry.Level != null) entryLevelDisplayName = ((StandardEventLevel) entry.Level).ToString();
                }

                string entryMessage = CleanUpMessage(entry.FormatDescription());
                if (entryMessage == null && entry.Properties.Count > 0)
                {
                    string entryProperties = String.Join(", ", entry.Properties.Select(p => String.Format("'{0}'", p.Value)));
                    entryMessage = String.Format(Coloma.descNotFoundMsgTemplate, entryId, entryProviderName, entryProperties);
                }

                if (entry.ProviderId.HasValue && entry.ProviderId.Value == servicingProvider)
                {
                    if (entryId == 2)
                    {
                        // this is a KB installed message, figure out which KB it is and update the revision of that entry
                        string kb = "KB";
                        int i = entryMessage.IndexOf(kb);

                        if (firstKbEvent == 0)
                        {
                            firstKbEvent = list.Count + 1;
                        }

                        if (-1 != i)
                        {
                            // we found the kb article
                            kb = entryMessage.Substring(i, 9);

                            foreach (KBRevision rev in Coloma.kbrlist)
                            {
                                if (rev.Kb == kb)
                                {
                                    Console.WriteLine(rev.Kb + " " + rev.Revision);
                                    revision = rev.Revision;
                                }
                            }
                        }
                    }
                }

                list.Add(new ColomaEvent(wvi.branch, wvi.build, revision, entryMachineName, DeviceId,
                    Environment.UserName, entryLogName, entryLevelDisplayName, entryId,
                    entryTimeCreated, entryProviderName, entryMessage));
            }

            // If there was a KB Installed event, go back and set all the events before it to have the default revision
            if (firstKbEvent != 0)
            {
                foreach (var evt in list)
                {
                    firstKbEvent--;
                    if (firstKbEvent <= 0)
                    {
                        break;
                    }

                    evt.Revision = Coloma.defaultRevision;
                }
            }
        }
        
        private static string CleanUpMessage(string message)
        {
            if (message != null)
            {
                message = message.Replace("\t", " ");
                message = message.Replace("\r\n", "<br>");
                message = message.Replace("\n", "<br>");
                message = message.Replace("\r", "<br>");
                message = message.Replace("<br><br>", "<br>");
            }

            return message;
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
            // First try CFC FLIGHTS, we have a backup key we can try if this doesn't work out.
            string keyName =
                "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Diagnostics\\DiagTrack\\SettingsRequests\\CFC.FLIGHTS";

            string valueName = @"ETagQueryParameters";

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

            // CFC FLIGHTS did not work, try SQMCLient.
            keyName = "SOFTWARE\\Microsoft\\SQMClient";
            valueName = @"MachineId";

            rk = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
            rk = rk.OpenSubKey(keyName);

            // We need to see if the key exists or contains what we need.
            if (rk != null)
            {
                string keyData = (string)rk.GetValue(valueName);

                // This key holds IDs in the form {XXXX-XXXX-XXXX-XXXX}
                // We need to remove the braces and add s: to the front.
                if (keyData.Contains("{") && keyData.Contains("}"))
                {
                    keyData = keyData.Replace("{", String.Empty);
                    keyData = keyData.Replace("}", String.Empty);

                    keyData = "s:" + keyData;
                    return keyData;
                }
            }

            // If we reach this point, deviceID is unknown
            return "s:UNKNOWN";
        }
    }
}