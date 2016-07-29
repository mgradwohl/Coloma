using System;
using System.Reflection;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Win32;

namespace Coloma
{
    public class Coloma
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

            // Parse args
            var onlyNewEvents = !args.Contains("-all");
            var saveLocal = args.Contains("-local");

            // Inform the user we're running
            Console.WriteLine();
            Console.WriteLine("Coloma is gathering log entries for your machine.");
            Console.WriteLine();

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

            // just get logs since last time OR since 4/1/2016
            DateTime dt = new DateTime(2016, 4, 1, 0, 0, 0, 0, DateTimeKind.Local);
            if (onlyNewEvents)
            {
                GetLastColomaDate(ref dt);
            }

            // Tell the user what we're doing
            Console.WriteLine("Any error, warning, or KB install written after " + dt.ToShortDateString());
            Console.WriteLine("From the following logs: System, Security, Hardware Events, Setup, and Application");
            Console.WriteLine("Data will be saved to " + filepath);

            if (DeviceId.Contains("UNKNOWN"))
            {
                Console.WriteLine("WARNING: Coloma could not automatically detect your DeviceID.");
            }

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
            uint? revision = GetRevisionFromSetupLog(kbrlist);
            Console.WriteLine("done");

            // Go through the 'standard' logs
            string[] Logs = { "System", "HardwareEvents", "Application", "Security", "Setup" };
            foreach (string log in Logs)
            {
                Console.Write(log + "... ");
                AddLogToList(log, eventlist, dt);
                Console.WriteLine("done");
            }

            // gets all the events in order, and uses info from the setup log to ensure the correct revision
            // if the setuplog had no entries, then use the current build and revision
            Console.Write("Sort and fixup... ");
            SortandFix(eventlist, revision);
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

        private static void SortandFix(List<ColomaEvent> eventlist, uint? revision)
        {
            uint r = 11;
            if (revision.HasValue)
            {
                r = revision.Value;
            }
            else
            {
                WindowsVersion.WindowsVersionInfo wvi = new WindowsVersion.WindowsVersionInfo();
                WindowsVersion.GetWindowsBuildandRevision(wvi);
                r = wvi.revision;
            }

            // sort the list by date
            eventlist.Sort();

            foreach (ColomaEvent evt in eventlist)
            {
                evt.Revision = r;
            }
        }

        private static uint? GetRevisionFromSetupLog(List<KBRevision> kbrlist)
        {
            EventLogQuery query = new EventLogQuery("Setup", PathType.LogName) {ReverseDirection = false};
            EventLogReader reader = new EventLogReader(query);

            EventRecord entry;
            while ((entry = reader.ReadEvent()) != null)
            {
                if (entry.Id != 2) continue;

                // this is a KB installed message, figure out which KB it is and update the revision of that entry
                string msg = CleanUpMessage(entry.FormatDescription());
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
                            return rev.Revision;
                        }
                    }
                }
            }

            return null;
        }

        private static void AddLogToList(string logName, List<ColomaEvent> list, DateTime dt)
        {
            // this retrieves the build.revision and branch for the current client
            WindowsVersion.WindowsVersionInfo wvi = new WindowsVersion.WindowsVersionInfo();
            WindowsVersion.GetWindowsBuildandRevision(wvi);

            EventLogQuery query = new EventLogQuery(logName, PathType.LogName);
            query.ReverseDirection = false;
            EventLogReader reader = new EventLogReader(query);

            EventRecord entry;
            while ((entry = reader.ReadEvent()) != null)
            {
                var entryTimeCreated = entry.TimeCreated ?? new DateTime();

                if (entryTimeCreated < dt)
                    continue;

                if ((entry.Level != (byte) StandardEventLevel.Critical) &&
                    (entry.Level != (byte) StandardEventLevel.Error) &&
                    (entry.Level != (byte) StandardEventLevel.Warning)) continue;

                var entryMachineName = entry.MachineName;
                var entryLogName = entry.LogName;
                var entryId = entry.Id;
                var entryProviderName = entry.ProviderName;
                var entryActivityId = entry.ActivityId;
                var entryProcessId = entry.ProcessId;
                var entryThreadId = entry.ThreadId;
                var entryProperties = entry.Properties;
                var entryQualifiers = entry.Qualifiers;
                var entryRecordId = entry.RecordId;
                var entryVersion = entry.Version;

                var entryPropertiesString = string.Join(", ", entryProperties.Select(
                    p => p.Value is byte[]
                        ? System.Text.Encoding.Unicode.GetString((byte[])p.Value).Replace(@"\0", "")
                        : p.Value));
                entryPropertiesString = CleanUpMessage(entryPropertiesString);

                string entryMessage = CleanUpMessage(entry.FormatDescription());
                if (entryMessage == null && entry.Properties.Count > 0)
                {
                    // Format the message to match how it is shown the event viewer
                    string descNotFoundMsgTemplate = "The description for Event ID '{0}' in Source '{1}' cannot be found.  The local computer may not have the necessary registry information or message DLL files to display the message, or you may not have permission to access them.  The following information is part of the event:{2}";
                    entryMessage = String.Format(descNotFoundMsgTemplate, entryId, entryProviderName, entryPropertiesString);
                }

                var entryUserId = "";
                try { entryUserId = entry.UserId.Value; }
                catch { }

                string entryLevelDisplayName = "";
                try { entryLevelDisplayName = entry.LevelDisplayName; }
                catch { if (entry.Level != null) entryLevelDisplayName = ((StandardEventLevel) entry.Level).ToString(); }

                var entryKeywords = entry.Keywords;
                var entryKeywordsDisplayNames = "";
                try { entryKeywordsDisplayNames = String.Join(", ", entry.KeywordsDisplayNames); }
                catch { entryKeywordsDisplayNames = entryKeywords.ToString(); }

                var entryOpcode = entry.Opcode;
                var entryOpcodeDisplayName = "";
                try { entryOpcodeDisplayName = entry.OpcodeDisplayName; }
                catch { entryOpcodeDisplayName = entryOpcode.ToString(); }

                var entryTask = entry.Task;
                var entryTaskDisplayName = "";
                try { entryTaskDisplayName = entry.TaskDisplayName; }
                catch { entryTaskDisplayName = entryTask.ToString(); }

                list.Add(new ColomaEvent(wvi.branch, wvi.build, 0, entryMachineName, DeviceId,
                    Environment.UserName, entryLogName, entryLevelDisplayName, entryId,
                    entryTimeCreated, entryProviderName, entryMessage, 
                    entryActivityId, entryProcessId, entryThreadId, entryPropertiesString, entryQualifiers, entryRecordId, 
                    entryUserId, entryVersion, entryKeywordsDisplayNames, entryOpcodeDisplayName, entryTaskDisplayName));
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