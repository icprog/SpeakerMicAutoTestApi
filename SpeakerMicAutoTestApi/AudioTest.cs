using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces;
using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SpeakerMicAutoTestApi
{
    public class AudioTest
    {
        PlatformFactory platformfactory = null;
        Platform platform = null;
        const string M101BProductName = "Agile_X";
        const string BartecProductName = "Agile_X_IS";

        public AudioTest()
        {
            Init();
        }

        void Init()
        {
            try
            {
                switch (GetPlatform())
                {
                    case M101BProductName:
                    case "M101B":
                        platform = new M101B();
                        Console.WriteLine("M101B");
                        break;
                    case BartecProductName:
                        platform = new Bartec();
                        Console.WriteLine("Bartec");
                        break;
                    default:
                        throw new Exception("Platform not support");
                        break;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                Console.WriteLine(ex);
            }

        }

        string GetPlatform()
        {
            string BIOSMainBoard = "";
            ManagementScope managementScope;
            ConnectionOptions connectionOptions;

            try
            {
                connectionOptions = new ConnectionOptions();
                connectionOptions.Impersonation = ImpersonationLevel.Impersonate;
                connectionOptions.Authentication = AuthenticationLevel.Default;
                connectionOptions.EnablePrivileges = true;

                managementScope = new ManagementScope();
                managementScope.Path = new ManagementPath(@"\\" + Environment.MachineName + @"\root\CIMV2");
                managementScope.Options = connectionOptions;

                SelectQuery selectQuery = new SelectQuery("SELECT * FROM Win32_ComputerSystemProduct");
                ManagementObjectSearcher managementObjectSearch = new ManagementObjectSearcher(managementScope, selectQuery);
                ManagementObjectCollection managementObjectCollection = managementObjectSearch.Get();

                foreach (ManagementObject managementObject in managementObjectCollection)
                {
                    BIOSMainBoard = (string)managementObject["Name"];
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                Console.WriteLine(ex);
            }

            return BIOSMainBoard;
        }


        public Platform.Result RunTest()
        {
            return platform.RunTest();
        }

        public Platform.Result AudioJackTest()
        {
            return platform.AudioJackTest();
        }

        public double InternalRecordThreshold
        {
            get { return platform.InternalRecordThreshold; }
            set { platform.InternalRecordThreshold = value; }
        }

        public double ExternalRecordThreshold
        {
            get { return platform.ExternalRecordThreshold; }
            set { platform.ExternalRecordThreshold = value; }
        }

        public double AudioJackRecordThreshold
        {
            get { return platform.AudioJackRecordThreshold; }
            set { platform.AudioJackRecordThreshold = value; }
        }

        public double LeftIntensity
        {
            get { return platform.LeftIntensity; }
        }

        public double RightIntensity
        {
            get { return platform.RightIntensity; }
        }

        public double InternalIntensity
        {
            get { return platform.InternalIntensity; }
        }

        public double AudioJackIntensity
        {
            get { return platform.AudioJackIntensity; }
        }

        public string WavFileName
        {
            get { return platform.WavFileName; }
            set { platform.WavFileName = value; }
        }
    }
}
