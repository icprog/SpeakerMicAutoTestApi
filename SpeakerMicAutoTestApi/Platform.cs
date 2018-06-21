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
using System.Threading.Tasks;

namespace SpeakerMicAutoTestApi
{
    public abstract class Platform
    {
        public enum Result
        {
            Pass,
            LeftSpeakerFail,
            RightSpeakerFail,
            InternalMicFail,
            InternalLeftMicFail,
            InternalRightMicFail,
            AudioJackFail,
            ExceptionFail
        }

        public enum Channel
        {
            Left,
            Right,
            HeadSet,
            AudioJack
        }

        protected double internalthreshold { get; set; }
        protected double externalthreshold { get; set; }
        protected double audiojackthreshold { get; set; }
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
        protected string ProductName { get; set; }
        protected int DeviceNumber { get; set; }
        protected float volume { get; set; }

        protected WaveInEvent WavSource = null;
        protected WaveFileWriter WavSourceFile = null;
        protected MMDeviceEnumerator DeviceEnum = null;
        protected Result result;
        public Exception exception;

        public Platform()
        {
            externalthreshold = 18000.0;
            internalthreshold = 18000.0;
            wavfilename = "o95.wav";
            LeftRecordFileName = "left.wav";
            RightRecordFileName = "right.wav";
            InternalRecordFileName = "headset.wav";
            LeftChannelFileName = "channel1.wav";
            RightChannelFileName = "channel2.wav";
            DeviceNumber = 0;
            ProductName = string.Empty;
            result = new Result();
            exception = null;
            leftintensity = 0.0;
            rightintensity = 0.0;
            internalintensity = 0.0;
            internalleftintensity = 0.0;
            internalrightintensity = 0.0;
            audiojackintensity = 0.0;
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
            set { wavfilename = value; }
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

        public abstract Result RunTest();
        public abstract Result AudioJackTest();
        protected abstract Task<string> Record(Channel Channel);
        protected abstract void PlayAndRecord(string WavFileName, Channel Channel);
        protected abstract double CalculateRMS(string WavFileName);
    }
}
