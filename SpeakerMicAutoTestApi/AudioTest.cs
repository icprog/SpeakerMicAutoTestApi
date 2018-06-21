using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces;
using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SpeakerMicAutoTestApi
{
    public class AudioTest
    {
        [DllImport(@"WMIO2.dll")]
        public static extern bool WinIO_ReadFromECSpace(uint uiAddress, out uint uiValue);

        Platform platform = null;
        const string M101BProductName = "IB80";
        const string BartecProductName = "BTZ1";

        public AudioTest()
        {
            Init();
        }

        void Init()
        {
            string Model = string.Empty;
            try
            {
                WinIO_GetECVersion(out Model);
                switch (Model.Substring(0,4))
                {
                    case M101BProductName:
                        platform = new M101B();
                        Console.WriteLine("M101B");
                        break;
                    case BartecProductName:
                        platform = new Bartec();
                        Console.WriteLine("Bartec");
                        break;
                    default:
                        throw new Exception("Platform not support");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                Console.WriteLine(ex);
            }
        }

        public static bool WinIO_GetECVersion(out string version)
        {
            uint bValue;
            version = string.Empty;

            WinIO_ReadFromECSpace(0x00, out bValue);
            if ((bValue < 0x20) || (bValue > 0x7A))
                bValue = 0x5F;
            version += Convert.ToChar(bValue).ToString();

            WinIO_ReadFromECSpace(0x01, out bValue);
            if ((bValue < 0x20) || (bValue > 0x7A))
                bValue = 0x5F;
            version += Convert.ToChar(bValue).ToString();

            WinIO_ReadFromECSpace(0x02, out bValue);
            if ((bValue < 0x20) || (bValue > 0x7A))
                bValue = 0x5F;
            version += Convert.ToChar(bValue).ToString();

            WinIO_ReadFromECSpace(0x03, out bValue);
            if ((bValue < 0x20) || (bValue > 0x7A))
                bValue = 0x20;
            version += Convert.ToChar(bValue).ToString();

            WinIO_ReadFromECSpace(0x04, out bValue);
            if ((bValue < 0x20) || (bValue > 0x7A))
                bValue = 0x20;
            version += Convert.ToChar(bValue).ToString();

            WinIO_ReadFromECSpace(0x05, out bValue);
            if ((bValue < 0x20) || (bValue > 0x7A))
                bValue = 0x20;
            version += Convert.ToChar(bValue).ToString();

            WinIO_ReadFromECSpace(0x06, out bValue);
            if ((bValue < 0x20) || (bValue > 0x7A))
                bValue = 0x20;
            version += Convert.ToChar(bValue).ToString();

            WinIO_ReadFromECSpace(0x07, out bValue);
            if ((bValue < 0x20) || (bValue > 0x7A))
                bValue = 0x20;
            version += Convert.ToChar(bValue).ToString();

            return true;
        }

        string GetBIOSPlatform()
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

        public Exception Exception
        {
            get { return platform.exception; }
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

        public double InternalLeftIntensity
        {
            get { return platform.InternalLeftIntensity; }
        }

        public double InternalRightIntensity
        {
            get { return platform.InternalRightIntensity; }
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
