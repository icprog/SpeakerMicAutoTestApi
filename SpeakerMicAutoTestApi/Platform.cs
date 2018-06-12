using NAudio.CoreAudioApi;
using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpeakerMicAutoTestApi
{
    public abstract class Platform
    {
        public enum Result
        {
            Pass = 0,
            LeftSpeakerFail = 1,
            RightSpeakerFail = 2,
            InternalMicFail = 3,
            AudioJackFail = 4,
            ExceptionFail = 5
        }

        protected double internalthreshold { get; set; }
        protected double externalthreshold { get; set; }
        protected double audiojackthreshold { get; set; }
        protected double audiojackintensity { get; set; }
        protected double leftintensity { get; set; }
        protected double rightintensity { get; set; }
        protected double internalintensity { get; set; }
        protected string wavfilename { get; set; }
        protected string LeftRecordFileName { get; set; }
        protected string RightRecordFileName { get; set; }
        protected string InternalRecordFileName { get; set; }
        protected string ProductName { get; set; }
        protected int DeviceNumber { get; set; }
        protected float volume { get; set; }

        protected WaveInEvent LeftSource = null;
        protected WaveFileWriter LeftSourceFile = null;
        protected WaveInEvent RightSource = null;
        protected WaveFileWriter RightSourceFile = null;
        protected WaveInEvent InternalSource = null;
        protected WaveFileWriter InternalSourceFile = null;
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
            DeviceNumber = 0;
            ProductName = string.Empty;
            result = new Result();
            exception = null;
            leftintensity = 10.0;
            rightintensity = 10.0;
            internalintensity = 10.0;
            audiojackintensity = 10.0;
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

        public abstract Result RunTest();
        public abstract Result AudioJackTest();
        protected abstract Task<string> RecordRightSpeaker();
        protected abstract Task<string> RecordLeftSpeaker();
        protected abstract Task<string> RecordHeadSet();
        protected abstract void PlayFromSpeakerAndRecord(string WavFileName, int channel);
        protected abstract void PlayFromHeadSetAndRecord(string WavFileName);
        protected abstract double CalculateRMS(string WavFileName);
    }
}
