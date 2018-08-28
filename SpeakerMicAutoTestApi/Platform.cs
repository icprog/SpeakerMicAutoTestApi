using IniParser;
using IniParser.Model;
using NAudio.CoreAudioApi;
using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Diagnostics;

namespace SpeakerMicAutoTestApi
{
    public class AudioDeviceComparer : IComparer<string>
    {
        string pattern = @"\(\D+.*";
        public int Compare(string x, string y)
        {
            if (string.Equals(x, y))
                return 0;

            if (Regex.IsMatch(x, pattern))
                return -1;

            if (Regex.IsMatch(y, pattern))
                return 1;

            return string.Compare(x, y, true);
        }
    }

    public abstract class Platform
    {
        [DllImport("MicrophoneBoost.dll", CallingConvention = CallingConvention.Cdecl)]
        public static extern int setMicrophoneBoost(float CurrentDb);

        public enum Result
        {
            Pass,
            LeftSpeakerFail,
            RightSpeakerFail,
            InternalMicFail,
            InternalLeftMicFail,
            InternalRightMicFail,
            AudioJackFail,
            FanRecordFail,
            ExceptionFail
        }

        public enum Channel
        {
            Left,
            Right,
            HeadSet,
            AudioJack,
            Fan
        }

        public enum AudioDeviceState : uint
        {
            Enable = 0x00000001,
            Disable = 0x10000001,
        }

        protected double internalthreshold { get; set; }
        protected double externalthreshold { get; set; }
        protected double audiojackthreshold { get; set; }
        protected double fanthreshold { get; set; }
        protected double fanintensity { get; set; }
        protected double audiojackintensity { get; set; }
        protected double leftintensity { get; set; }
        protected double rightintensity { get; set; }
        protected double internalintensity { get; set; }
        protected double internalleftintensity { get; set; }
        protected double internalrightintensity { get; set; }
        protected string wavfilename { get; set; }
        protected string LeftRecordFileName { get; set; }
        protected string RightRecordFileName { get; set; }
        protected string InternalRecordFileName { get; set; }
        protected string LeftChannelFileName { get; set; }
        protected string RightChannelFileName { get; set; }
        protected string FanRecordFileName { get; set; }
        protected string AudioJackRecordFileName { get; set; }
        protected string RealtekMicrophone { get; set; }
        protected string PinkMicrophone { get; set; }
        protected string ProductName { get; set; }
        protected int DeviceNumber { get; set; }
        protected float volume { get; set; }
        protected int AudioTimeout { get; set; }
        protected bool OriginalState { get; set; }
        protected bool Success = false;

        protected WaveInEvent WavSource = null;
        protected WaveFileWriter WavSourceFile = null;
        protected MMDeviceEnumerator DeviceEnum = null;
        protected Result result;
        public Exception exception;

        public Platform()
        {
            externalthreshold = 18000.0;
            internalthreshold = 18000.0;
            audiojackthreshold = 18000.0;
            fanthreshold = 18000.0;
            AudioTimeout = 10000;
            wavfilename = GetFullPath("WinmateAudioTest.wav");
            LeftRecordFileName = GetFullPath("left.wav");
            RightRecordFileName = GetFullPath("right.wav");
            InternalRecordFileName = GetFullPath("headset.wav");
            LeftChannelFileName = GetFullPath("channel1.wav");
            RightChannelFileName = GetFullPath("channel2.wav");
            AudioJackRecordFileName = GetFullPath("audiojack.wav");
            FanRecordFileName = GetFullPath("fan.wav");
            DeviceNumber = 0;
            ProductName = string.Empty;
            result = new Result();
            exception = null;
            leftintensity = -1;
            rightintensity = -1;
            internalintensity = -1;
            internalleftintensity = -1;
            internalrightintensity = -1;
            audiojackintensity = -1;
            fanintensity = -1;
            RealtekMicrophone = @"(Microphone \(\d*-*\s*Realtek High)";
            PinkMicrophone = @"(Mic in at rear panel \(Pink\) \(\d*-*\s*Re)";
            OriginalState = false;
        }

        public void DeleteRecordWav()
        {
            if (File.Exists(LeftRecordFileName)) File.Delete(LeftRecordFileName);
            if (File.Exists(RightRecordFileName)) File.Delete(RightRecordFileName);
            if (File.Exists(InternalRecordFileName)) File.Delete(InternalRecordFileName);
            if (File.Exists(LeftChannelFileName)) File.Delete(LeftChannelFileName);
            if (File.Exists(RightChannelFileName)) File.Delete(RightChannelFileName);
            if (File.Exists(AudioJackRecordFileName)) File.Delete(AudioJackRecordFileName);
            if (File.Exists(FanRecordFileName)) File.Delete(FanRecordFileName);
        }

        public double FanRecordThreshold
        {
            get { return fanthreshold; }
            set { fanthreshold = value; }
        }

        public double AudioJackRecordThreshold
        {
            get { return audiojackthreshold; }
            set { audiojackthreshold = value; }
        }

        public double InternalRecordThreshold
        {
            get { return internalthreshold; }
            set { internalthreshold = value; }
        }

        public double ExternalRecordThreshold
        {
            get { return externalthreshold; }
            set { externalthreshold = value; }
        }

        public double FanIntensity
        {
            get { return fanintensity; }
        }

        public double AudioJackIntensity
        {
            get { return audiojackintensity; }
        }

        public double LeftIntensity
        {
            get { return leftintensity; }
        }

        public double RightIntensity
        {
            get { return rightintensity; }
        }

        public double InternalIntensity
        {
            get { return internalintensity; }
        }

        public double InternalLeftIntensity
        {
            get { return internalleftintensity; }
        }

        public double InternalRightIntensity
        {
            get { return internalrightintensity; }
        }

        public string WavFileName
        {
            get { return wavfilename; }
            set { wavfilename = GetFullPath(value); }
        }

        public int LeftVolume
        {
            set
            {
                if (value > 100)
                    value = 100;

                DeviceEnum = new MMDeviceEnumerator();
                var collect = DeviceEnum.EnumerateAudioEndPoints(DataFlow.All, DeviceState.Active);
                foreach (var v in collect)
                {
                    if (v.DataFlow == DataFlow.Capture)
                        v.AudioEndpointVolume.MasterVolumeLevelScalar = 100 / 100f;

                    if (v.DataFlow == DataFlow.Render)
                        v.AudioEndpointVolume.Channels[0].VolumeLevelScalar = value / 100f;
                }
            }
        }

        public int RightVolume
        {
            set
            {
                if (value > 100)
                    value = 100;

                DeviceEnum = new MMDeviceEnumerator();
                var collect = DeviceEnum.EnumerateAudioEndPoints(DataFlow.All, DeviceState.Active);
                foreach (var v in collect)
                {
                    if (v.DataFlow == DataFlow.Capture)
                        v.AudioEndpointVolume.MasterVolumeLevelScalar = 100 / 100f;

                    if (v.DataFlow == DataFlow.Render)
                        v.AudioEndpointVolume.Channels[1].VolumeLevelScalar = value / 100f;
                }
            }
        }

        public float MicrophoneBoost
        {
            set
            {
                if (value > 30.0f)
                    value = 30.0f;

                setMicrophoneBoost(value);
            }
        }

        protected string GetIniValue(string SectionName, string Key)
        {
            var parser = new FileIniDataParser();
            parser.Parser.Configuration.CommentString = ";";
            var FilePath = Assembly.GetEntryAssembly().Location;
            IniData data = parser.ReadFile(Path.Combine(Directory.GetParent(FilePath).FullName, "TestProgramConfig.ini"));
            var FormatKey = string.Format("{0}.{1}", SectionName, Key);
            Console.WriteLine(data.GetKey(FormatKey));
            return data.GetKey(FormatKey);
        }

        protected string GetConfigValue(string Key)
        {
            JObject JObject = JObject.Parse(File.ReadAllText(GetFullPath("config.json")));
            return JObject[Key].ToString();
        }

        protected string GetFullPath(string path)
        {
            var exepath = System.AppDomain.CurrentDomain.BaseDirectory;
            Console.WriteLine(Path.Combine(exepath, path));
            return Path.Combine(exepath, path);
        }

        private static void ShowSecurity(RegistrySecurity security)
        {
            Console.WriteLine("\r\nCurrent access rules:\r\n");

            foreach (RegistryAccessRule ar in
                security.GetAccessRules(true, true, typeof(NTAccount)))
            {
                Console.WriteLine("        User: {0}", ar.IdentityReference);
                Console.WriteLine("        Type: {0}", ar.AccessControlType);
                Console.WriteLine("      Rights: {0}", ar.RegistryRights);
                Console.WriteLine();
            }
        }

        public void TakeOwnership(string key)
        {
            try
            {
                TokenManipulator.AddPrivilege("SeRestorePrivilege");
                TokenManipulator.AddPrivilege("SeBackupPrivilege");
                TokenManipulator.AddPrivilege("SeTakeOwnershipPrivilege");

                SecurityIdentifier sid = new SecurityIdentifier(WellKnownSidType.BuiltinUsersSid, null);
                NTAccount account = sid.Translate(typeof(NTAccount)) as NTAccount;
                using (var HKLM = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
                using (var Registry = HKLM.OpenSubKey(key, RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryRights.TakeOwnership))
                {
                    var rs = Registry.GetAccessControl(AccessControlSections.All);
                    rs.SetOwner(account);
                    Registry.SetAccessControl(rs);
                }
            }
            finally
            {
                TokenManipulator.RemovePrivilege("SeRestorePrivilege");
                TokenManipulator.RemovePrivilege("SeBackupPrivilege");
                TokenManipulator.RemovePrivilege("SeTakeOwnershipPrivilege");
            }
        }

        public void RemoveProtection(string key)
        {
            try
            {
                TokenManipulator.AddPrivilege("SeRestorePrivilege");
                TokenManipulator.AddPrivilege("SeBackupPrivilege");
                TokenManipulator.AddPrivilege("SeTakeOwnershipPrivilege");

                using (var HKLM = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
                using (var Registry = HKLM.OpenSubKey(key, RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryRights.ChangePermissions))
                {
                    var rs = Registry.GetAccessControl(AccessControlSections.All);
                    rs.SetAccessRuleProtection(true, true);
                    rs.AddAccessRule(new RegistryAccessRule("Users", RegistryRights.FullControl, AccessControlType.Allow));
                    Registry.SetAccessControl(rs);
                }
            }
            finally
            {
                TokenManipulator.RemovePrivilege("SeRestorePrivilege");
                TokenManipulator.RemovePrivilege("SeBackupPrivilege");
                TokenManipulator.RemovePrivilege("SeTakeOwnershipPrivilege");
            }
        }

        public string enumerateKeys(string keyPath, string name)
        {
            try
            {
                TakeOwnership(keyPath);
                RemoveProtection(keyPath);

                using (var HKLM = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
                using (var RegKey = HKLM.OpenSubKey(keyPath))
                {
                    var subKeys = RegKey.GetSubKeyNames();
                    var subNames = RegKey.GetValueNames();
                    var MicrophoneFlag = false;
                    var RealtekFlag = false;
                    var test = @"High Definition Audio Device";
                    var Microphone = @"Microphone";
                    var keypath = string.Empty;

                    foreach (var v in subNames)
                    {
                        if (RegKey.GetValue(v) is string)
                        {
                            if (RegKey.GetValue(v).ToString().Contains(test))
                                RealtekFlag = true;

                            if (RegKey.GetValue(v).ToString().Contains(Microphone))
                                MicrophoneFlag = true;

                            if (RealtekFlag && MicrophoneFlag)
                            {
                                Console.WriteLine(Directory.GetParent(RegKey.ToString()));
                                keypath = Directory.GetParent(RegKey.ToString()).ToString();
                                break;
                            }
                        }
                    }

                    foreach (string subKey in subKeys)
                    {
                        string fullPath = keyPath + "\\" + subKey;
                        enumerateKeys(fullPath, name);
                    }

                    return keypath;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.WriteLine("error path: {0}", keyPath);
                return keyPath;
            }
            finally
            {
            }
        }

        public bool SetAudioDeviceState(string Name, bool Enable, out bool Success, DeviceState state)
        {
            string test = "Microphone (High Definition Audio Device)";
            bool IsEnable = false;
            Success = false;
            Regex regex = new Regex(Name);

            try
            {
                DeviceEnum = new MMDeviceEnumerator();
                var collect = DeviceEnum.EnumerateAudioEndPoints(DataFlow.Capture, state);

                var device = collect.Where(e => !string.IsNullOrEmpty(e.FriendlyName) && regex.IsMatch(e.FriendlyName));

                if (!device.Any())
                {
                    Console.WriteLine("{0} not found", Name);
                    return true;
                }

                foreach (var v in device)
                {
                    var guid = v.ID.Split(new string[] { "{", "}", "." }, StringSplitOptions.RemoveEmptyEntries);
                    Console.WriteLine(v.FriendlyName);
                    Console.WriteLine(v.ID);
                    Console.WriteLine(guid.LastOrDefault());
#if true
                    var subkey = string.Format(@"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture\{{{0}}}", guid.LastOrDefault());
                    TakeOwnership(subkey);
                    RemoveProtection(subkey);

                    using (var hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
                    using (var key = hklm.OpenSubKey(subkey, true))
                    {
                        Console.WriteLine("DeviceState: {0}", key.GetValue("DeviceState"));
                        if (Convert.ToUInt32(key.GetValue("DeviceState")) != Convert.ToUInt32(AudioDeviceState.Enable)
                            && Convert.ToUInt32(key.GetValue("DeviceState")) != Convert.ToUInt32(AudioDeviceState.Disable))
                            continue;

                        IsEnable = Convert.ToUInt32(key.GetValue("DeviceState")) == Convert.ToUInt32(AudioDeviceState.Enable) ? true : false;

                        if (Enable)
                            key.SetValue("DeviceState", AudioDeviceState.Enable, RegistryValueKind.DWord);
                        else
                            key.SetValue("DeviceState", AudioDeviceState.Disable, RegistryValueKind.DWord);

                        Console.WriteLine(IsEnable);
                    }
#endif
                }
                Success = true;
                return IsEnable;
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex);
                Console.WriteLine(ex);
                exception = ex;
                Success = false;
                return IsEnable;
            }
            finally
            {

            }
        }

        public abstract Result RunTest();
        public abstract Result AudioJackTest();
        public abstract Result FanTest();
        protected abstract Task<string> Record(Channel Channel);
        protected abstract void PlayAndRecord(string WavFileName, Channel Channel);
        protected abstract double CalculateRMS(string WavFileName);
    }
}
