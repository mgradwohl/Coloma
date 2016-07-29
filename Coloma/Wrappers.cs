using Microsoft.Win32;
using System;
using System.Linq;

namespace Coloma
{
    public class ColomaEvent : IComparable
    {
        private string branch;
        private string level;
        private string machineName;
        private string deviceId;
        private string userName;
        private long instanceid;
        private DateTime timeCreated;
        private string source;
        private string message;

        private Guid? entryActivityId;
        private int? entryProcessId;
        private int? entryThreadId;
        private string entryProperties;
        private int? entryQualifiers;
        private long? entryRecordId;
        private string entryUserId;
        private byte? entryVersion;
        private string entryKeywordsDisplayNames;
        private string entryOpcodeDisplayName;
        private string entryTaskDisplayName;

        public uint Revision { get; set; }

        public string Logname { get; set; }

        public uint Build { get; set; }

        public ColomaEvent( string branch,
                            uint build,
                            uint revision,
                            string machineName,
                            string deviceId,
                            string userName,
                            string logName,
                            string level,
                            long id,
                            DateTime timeCreated,
                            string source,
                            string Message,

                            Guid? entryActivityId,
                            int? entryProcessId,
                            int? entryThreadId,
                            string entryProperties,
                            int? entryQualifiers,
                            long? entryRecordId,
                            string entryUserId,
                            byte? entryVersion,
                            string entryKeywordsDisplayNames,
                            string entryOpcodeDisplayName,
                            string entryTaskDisplayName)
        {
            this.branch = branch;
            this.Build = build;
            this.Revision = revision;
            this.machineName = machineName;
            this.deviceId = deviceId;
            this.userName = userName;
            this.Logname = logName;
            this.level = level;
            this.timeCreated = timeCreated;
            this.source = source;
            this.message = Message;
            this.instanceid = id;

            this.entryActivityId = entryActivityId;
            this.entryProcessId = entryProcessId;
            this.entryThreadId = entryThreadId;
            this.entryProperties = entryProperties;
            this.entryQualifiers = entryQualifiers;
            this.entryRecordId = entryRecordId;
            this.entryUserId = entryUserId;
            this.entryVersion = entryVersion;
            this.entryKeywordsDisplayNames = entryKeywordsDisplayNames;
            this.entryOpcodeDisplayName = entryOpcodeDisplayName;
            this.entryTaskDisplayName = entryTaskDisplayName;
        }

        public int CompareTo(object obj)
        {
            if (obj == null) return 1;

            ColomaEvent eventToCompare = obj as ColomaEvent;
            if (eventToCompare.timeCreated < timeCreated)
            {
                return 1;
            }
            if (eventToCompare.timeCreated > timeCreated)
            {
                return -1;
            }

            return 0;
        }

        public override string ToString()
        {
            string ret = string.Join("\t", branch, Build.ToString(), Revision.ToString(), machineName, deviceId,
                                     userName, Logname, level, instanceid.ToString(), timeCreated.ToString(), source, message, 
                                     entryActivityId, entryProcessId, entryThreadId, entryProperties, entryQualifiers, entryRecordId, 
                                     entryUserId, entryVersion, entryKeywordsDisplayNames, entryOpcodeDisplayName, entryTaskDisplayName);
            return ret;
        }

        public static string Header()
        {
            string ret = string.Join("\t", nameof(branch), nameof(Build), nameof(Revision), nameof(machineName), nameof(deviceId),
                                     nameof(userName), nameof(Logname), nameof(level), nameof(instanceid), nameof(timeCreated), nameof(source), nameof(message),
                                     nameof(entryActivityId), nameof(entryProcessId), nameof(entryThreadId), nameof(entryProperties), nameof(entryQualifiers), nameof(entryRecordId),
                                     nameof(entryUserId), nameof(entryVersion), nameof(entryKeywordsDisplayNames), nameof(entryOpcodeDisplayName), nameof(entryTaskDisplayName));
            return ret;
        }

    }

    public class KBRevision
    {
        public uint Build { get; set; }

        public uint Revision { get; set; }

        public string Kb { get; set; }

        public DateTime Installdate { get; set; }

        public KBRevision(uint b, uint r, string k)
        {
            Build = b;
            Revision = r;
            Kb = k;
            Installdate = new DateTime();
        }
    }

    public class WindowsVersion
    {
        // https://osgwiki.com/wiki/OS_Versioning_Identifiers_for_Servicing
        // HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\BuildBranch
        // HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\UBR
        // HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentBuild

        //BuildBranch = the physical branch name that built the image.e.g. for TH RTM this is "TH1". It does not change until the OS is upgraded to a new OS.For the released TH2, this is TH2_RELEASE
        //CurrentBuild = Build number that created the image e.g.TH1 RTM is 10240. It does not change until the OS is upgraded to a new OS
        //UBR - AKA "Revision" or "QFE" Number.This is initially set to the revision number of the RTM build.It is the last part of a full version string e.g. 10.0.10240.16384 (16384 = revision number).
        //  Its value changes when update packages are installed(this value is defined in source as VER_PRODUCTBUILD_QFE). 

        public class WindowsVersionInfo
        {
            public string branch;
            public uint build;
            public uint revision;
        }

        public static void GetWindowsBuildandRevision(WindowsVersionInfo wvi)
        {
            const string keyName = "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion";

            //get the 64-bit view first - if you read the 32 bit keys on a 64 bit machine you don't get the revision
            RegistryKey rk = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, RegistryView.Registry64);
            rk = rk.OpenSubKey(keyName);

            // 64bit registry didn't open, open the 32bit registry
            if (rk == null)
            {
                //we couldn't find the value in the 64-bit view so grab the 32-bit view
                rk = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, RegistryView.Registry32);
                rk = rk.OpenSubKey(keyName);
            }

            if (rk != null)
            {
                wvi.branch = (string)rk.GetValue("BuildBranch");
                // all these converts suck
                wvi.build = Convert.ToUInt32(rk.GetValue("CurrentBuild").ToString());
                wvi.revision = Convert.ToUInt32(rk.GetValue("UBR").ToString());
            }
            rk.Close();
        }
    }
}