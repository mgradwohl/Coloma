using System;
using System.Reflection;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Management.Infrastructure;
using Microsoft.Win32;
using System.Diagnostics;
using System.Globalization;

namespace Coloma
{
    public class Coloma
    {
        // Just load deviceId once, it's independant of other operations.
        private static readonly string DeviceId = GetDeviceId();

        private static readonly string descNotFoundMsgTemplate = "The description for Event ID '{0}' in Source '{1}' cannot be found.  The local computer may not have the necessary registry information or message DLL files to display the message, or you may not have permission to access them.  The following information is part of the event:{2}";
        private static readonly string structuredQueryTemplate = "<QueryList> <Query Id=\"0\"> <Select Path=\"Setup\">*[System[(Level=1 or Level=2 or Level=3) or (EventID=2)]]</Select> <Select Path=\"Application\">*[System[(Level=1 or Level=2 or Level=3) and TimeCreated[@SystemTime&gt;='{0:O}']]]</Select> <Select Path=\"System\">*[System[((Level=1 or Level=2 or Level=3) and TimeCreated[@SystemTime&gt;='{0:O}']) or ((EventID=6009) and Provider[@Name='EventLog'])]]</Select> <Select Path=\"Security\">*[System[(Level=1 or Level=2 or Level=3) and TimeCreated[@SystemTime&gt;='{0:O}']]]</Select> <Select Path=\"HardwareEvents\">*[System[(Level=1 or Level=2 or Level=3) and TimeCreated[@SystemTime&gt;='{0:O}']]]</Select> </Query> </QueryList>";

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

            Console.WriteLine("Querying installed updates...");
            var installedUpdates = QueryUpdates();

            List<ColomaEvent> eventlist = new List<ColomaEvent>();

            Console.WriteLine("Collecting events...");

            Stopwatch timer = Stopwatch.StartNew();
            AddLogToList(eventlist, dt, installedUpdates);
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

        struct FirstBuild
        {
            public int index;
            public uint build;
        }

        private static void AddLogToList(List<ColomaEvent> list, DateTime dt, Dictionary<string, KBRevision> installedUpdates)
        {
            // this retrieves the build.revision and branch for the current client
            WindowsVersion.WindowsVersionInfo wvi = new WindowsVersion.WindowsVersionInfo();
            WindowsVersion.GetWindowsBuildandRevision(wvi);

            FirstBuild firstBuildEvent = new FirstBuild { index = 0, build = wvi.build };

            KBRevision currentKb = new KBRevision();
            uint currentBuild = wvi.build;

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
                var entryTimeCreated = entry.TimeCreated ?? DateTime.Now;
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

                if (entryId == 2 && entry.ProviderId.HasValue && entry.ProviderId.Value == servicingProvider)
                {
                    if (entry.Properties.Count == 5 && entry.Properties[2].Value.ToString() == "Installed")
                    {
                        string kb = entry.Properties[0].Value.ToString();

                        KBRevision KBRev;
                        if(installedUpdates.TryGetValue(kb, out KBRev))
                        {
                            // Update InstallDate with a more precise time
                            KBRev.InstallDate = entryTimeCreated;
                            KBRev.FoundInSetupLog = true;

                            currentKb = KBRev;
                        }
                        else
                        {
                            currentKb = new KBRevision { InstallDate = entryTimeCreated, Kb = kb, FoundInSetupLog = true };
                            installedUpdates.Add(kb, new KBRevision(currentKb));
                        }
                    }
                }
                else if (entryId == 6009 && entry.ProviderName == "EventLog")
                {
                    if (firstBuildEvent.index == 0 && entry.Properties.Count == 5)
                    {
                        uint.TryParse(entry.Properties[1].Value.ToString(), out currentBuild);

                        // If the boot event indicates an earlier build, remember it so we can set all the prior events to this build number
                        if (firstBuildEvent.build != currentBuild)
                        {
                            firstBuildEvent.index = list.Count + 1;
                            firstBuildEvent.build = currentBuild;
                        }
                        // Else the event indicates no build number change, so don't bother checking later boot events either
                        else
                        {
                            firstBuildEvent.index = -1;
                        }
                    }

                    // This event doesn't go into the log file
                    continue;
                }

                list.Add(new ColomaEvent(wvi.branch, currentBuild, currentKb, entryMachineName, DeviceId,
                    Environment.UserName, entryLogName, entryLevelDisplayName, entryId,
                    entryTimeCreated, entryProviderName, entryMessage));
            }

            IEnumerable<KBRevision> orderedUpdates = installedUpdates.Select(kb => kb.Value).OrderBy(kb => kb.InstallDate);
            IEnumerable<KBRevision> missingUpdates = orderedUpdates.Where(kb => !kb.FoundInSetupLog);
            IEnumerator<KBRevision> currentMissingEnum = missingUpdates.GetEnumerator(); currentMissingEnum.MoveNext();
            IEnumerator<KBRevision> nextInstalled = orderedUpdates.GetEnumerator(); nextInstalled.MoveNext();
            DateTime nextInstallTime = DateTime.Now;

            Action getNextInstallTime = new Action(() =>
            {
                if (nextInstalled.Current != null && currentMissingEnum != null && currentMissingEnum.Current != null)
                {
                    do
                    {
                        if (nextInstalled.Current.Kb == currentMissingEnum.Current.Kb)
                        {
                            if (nextInstalled.MoveNext())
                            {
                                nextInstallTime = nextInstalled.Current.InstallDate;

                                // Install times from WMI only include the date, so skip ahead to the next install that falls on a later date or has a more precise time from a setup event
                                while (!nextInstalled.Current.FoundInSetupLog && nextInstallTime >= currentMissingEnum.Current.InstallDate && nextInstallTime < currentMissingEnum.Current.InstallDate.AddDays(1))
                                {
                                    if (nextInstalled.MoveNext())
                                    {
                                        nextInstallTime = nextInstalled.Current.InstallDate;
                                    }
                                    else
                                    {
                                        nextInstallTime = DateTime.Now;
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                nextInstallTime = DateTime.Now;
                                break;
                            }

                            break;
                        }
                    }
                    while (nextInstalled.MoveNext());
                }
            });

            getNextInstallTime();

            int fixupsRemaining = (firstBuildEvent.index > 0 ? 1 : 0) + (missingUpdates.Any() ? 1 : 0);

            foreach (var evt in list)
            {
                if (fixupsRemaining == 0)
                {
                    break;
                }

                // If there was a EventLog 6009 event with a different build number, go back and set all the events before it to have that build number
                // This is highly unlikely to ever happen under the current OS-swap upgrade mechanism because the logs get cleared during each upgrade
                if (firstBuildEvent.index > 0)
                {
                    firstBuildEvent.index--;
                    if (firstBuildEvent.index <= 0)
                    {
                        fixupsRemaining--;
                    }
                    else
                    {
                        evt.Build = firstBuildEvent.build;
                    }
                }

                // If there are installed KBs found via WMI that were not in the Setup log, go back and set all the events between the KB installation time
                // and the next KB installation time to use this KB number.
                if (currentMissingEnum != null && currentMissingEnum.Current != null)
                {
                    if (evt.TimeCreated >= currentMissingEnum.Current.InstallDate && evt.TimeCreated < nextInstallTime)
                    {
                        evt.MostRecentKb = new KBRevision(currentMissingEnum.Current);

                        // Since the installations times from WMI are only accurate to the day, append a * to the KB name for any events that fall on that day, to indicate uncertainty.
                        if (evt.TimeCreated >= evt.MostRecentKb.InstallDate && evt.TimeCreated < evt.MostRecentKb.InstallDate.AddDays(1))
                        {
                            evt.MostRecentKb.Kb += "*";
                        }
                    }
                    else if (evt.TimeCreated >= nextInstallTime)
                    {
                        if (currentMissingEnum.MoveNext())
                        {
                            getNextInstallTime();
                        }
                        else
                        {
                            currentMissingEnum = null;
                            fixupsRemaining--;
                        }
                    }
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

        // Attempts to retrieve the device ID from the registry
        //
        // Note: The device ID here is a token retrieved by UTC from a web service.
        //       Only Vortex is able to convert this token into the global device ID used in the Cosmos telemetry streams
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

        private static Dictionary<string, KBRevision> QueryUpdates()
        {
            Dictionary<string, KBRevision> installedUpdates = new Dictionary<string, KBRevision>();

            try
            {
                // TODO: Use a different WMI client that doesn't require WinRM

                CimSession cimSession = CimSession.Create("localhost");
                IEnumerable<CimInstance> enumeratedInstances = cimSession.EnumerateInstances(@"root\cimv2", "Win32_QuickFixEngineering");
                foreach (CimInstance cimInstance in enumeratedInstances)
                {
                    DateTime date;
                    if (!DateTime.TryParse(cimInstance.CimInstanceProperties["InstalledOn"].Value.ToString(), CultureInfo.CurrentCulture.DateTimeFormat, DateTimeStyles.AssumeUniversal, out date))
                    {
                        long time;
                        if (!long.TryParse("0x" + cimInstance.CimInstanceProperties["InstalledOn"].Value.ToString(), out time)) // This may only be needed for Win7...
                        {
                            continue;
                        }

                        date = DateTime.FromFileTimeUtc(time);
                    }

                    KBRevision rev = new KBRevision();
                    rev.Kb = cimInstance.CimInstanceProperties["HotFixID"].Value.ToString();
                    rev.InstallDate = date;

                    installedUpdates[rev.Kb] = rev;
                }
            }
            catch (Exception)
            {
                // This probably means WinRM is not enabled. The MI .Net interface uses WinRM to communicate with WMI
                // The user would need to run "winrm quickconfig" to setup the winrm service and firewall exceptions
            }

            return installedUpdates;
        }
    }
}
