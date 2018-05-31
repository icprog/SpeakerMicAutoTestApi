using NAudio.CoreAudioApi;
using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SpeakerMicAutoTestApi
{
    class CustomException : Exception
    {
        public CustomException(string message)
        {

        }

    }

    public enum Result
    {
        Pass = 0,
        LeftSpeakerFail = 1,
        RightSpeakerFail = 2,
        InternalMicFail = 3,
        ExceptionFail = 4
    }

    public class AudioTest
    {
        double internalthreshold { get; set; }
        double externalthreshold { get; set; }
        double leftintensity { get; set; }
        double rightintensity { get; set; }
        double internalintensity { get; set; }
        string wavfilename { get; set; }
        string LeftRecordFileName { get; set; }
        string RightRecordFileName { get; set; }
        string InternalRecordFileName { get; set; }
        string ProductName { get; set; }
        int DeviceNumber { get; set; }
        float volume { get; set; }

        WaveInEvent LeftSource = null;
        WaveFileWriter LeftSourceFile = null;
        WaveInEvent RightSource = null;
        WaveFileWriter RightSourceFile = null;
        WaveInEvent InternalSource = null;
        WaveFileWriter InternalSourceFile = null;
        AutoResetEvent[] RecordEvent = null;
        MMDeviceEnumerator DeviceEnum = null;
        Result result;

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

        public AudioTest()
        {
            externalthreshold = 6000.0;
            internalthreshold = 20000.0;
            wavfilename = "o95.wav";
            LeftRecordFileName = "left.wav";
            RightRecordFileName = "right.wav";
            InternalRecordFileName = "headset.wav";
            DeviceNumber = 0;
            ProductName = string.Empty;
            RecordEvent = new AutoResetEvent[3] { new AutoResetEvent(false), new AutoResetEvent(false), new AutoResetEvent(false) };
            result = new Result();
            leftintensity = 0.0;
            rightintensity = 0.0;
            internalintensity = 0.0;
        }

        public Result RunTest()
        {
            try
            {
                result = Result.Pass;
                var left = Task.Factory.StartNew(() =>
                {
                    LeftVolume = 100;
                    RightVolume = 0;
                    PlayFromSpeakerAndRecord(WavFileName);
                });

                left.Wait();
                
                //Task.Factory.StartNew(() =>
                //{
                //    LeftVolume = 100;
                //    RightVolume = 0;
                //    PlayFromSpeakerAndRecord(WavFileName);
                //});
                //Console.WriteLine("1");
                //RecordEvent[0].WaitOne();
                //RecordEvent[1].WaitOne();
                //Console.WriteLine("2");

                leftintensity = CalculateRMS(LeftRecordFileName);
                if (leftintensity < externalthreshold)
                    result = Result.LeftSpeakerFail;

                var right = Task.Factory.StartNew(() =>
                {
                    LeftVolume = 0;
                    RightVolume = 100;
                    PlayFromSpeakerAndRecord(WavFileName);
                });

                right.Wait();
                //RecordEvent[0].WaitOne();
                //RecordEvent[1].WaitOne();

                rightintensity = CalculateRMS(RightRecordFileName);
                if (rightintensity < externalthreshold)
                    result = Result.RightSpeakerFail;

                var headset = Task.Factory.StartNew(() =>
                {
                    LeftVolume = 100;
                    RightVolume = 100;
                    PlayFromHeadSetAndRecord(WavFileName);
                });

                headset.Wait();
                //RecordEvent[2].WaitOne();

                internalintensity = CalculateRMS(InternalRecordFileName);
                if (CalculateRMS(InternalRecordFileName) < internalthreshold)
                    result = Result.InternalMicFail;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                result = Result.ExceptionFail;
            }

            return result;
        }

        void InternalSource_DataAvailable(object sender, WaveInEventArgs e)
        {
            if (InternalSourceFile != null)
            {
                InternalSourceFile.Write(e.Buffer, 0, e.BytesRecorded);
                InternalSourceFile.Flush();
            }
        }

        void InternalSource_RecordingStopped(object sender, StoppedEventArgs e)
        {
            if (InternalSource != null)
            {
                InternalSource.Dispose();
                InternalSource = null;
            }

            if (InternalSourceFile != null)
            {
                InternalSourceFile.Dispose();
                InternalSourceFile = null;
            }

            Thread.Sleep(200);
            RecordEvent[2].Set();
        }

        void RightSource_DataAvailable(object sender, WaveInEventArgs e)
        {
            if (RightSourceFile != null)
            {
                RightSourceFile.Write(e.Buffer, 0, e.BytesRecorded);
                RightSourceFile.Flush();
            }
        }

        void RightSource_RecordingStopped(object sender, StoppedEventArgs e)
        {
            if (RightSource != null)
            {
                RightSource.Dispose();
                RightSource = null;
            }

            if (RightSourceFile != null)
            {
                RightSourceFile.Dispose();
                RightSourceFile = null;
            }

            Thread.Sleep(200);
            RecordEvent[1].Set();
        }

        void LeftSource_DataAvailable(object sender, WaveInEventArgs e)
        {
            if (LeftSourceFile != null)
            {
                LeftSourceFile.Write(e.Buffer, 0, e.BytesRecorded);
                LeftSourceFile.Flush();
            }
        }

        void LeftSource_RecordingStopped(object sender, StoppedEventArgs e)
        {
            if (LeftSource != null)
            {
                LeftSource.Dispose();
                LeftSource = null;
            }

            if (LeftSourceFile != null)
            {
                LeftSourceFile.Dispose();
                LeftSourceFile = null;
            }

            Thread.Sleep(200);
            RecordEvent[0].Set();
        }

        void RecordRightSpeaker()
        {
            //try
            {
                RightSource = new WaveInEvent();
                string AudioPrefix = "Logitech USB H";
                ProductName = string.Empty;

                for (int n = -1; n < WaveInEvent.DeviceCount; n++)
                {
                    var caps = WaveInEvent.GetCapabilities(n);
                    Console.WriteLine("Record device {0}: {1}", n, caps.ProductName);
                    if (caps.ProductName.Contains(AudioPrefix))
                    {
                        DeviceNumber = n;
                        ProductName = caps.ProductName;
                        Console.WriteLine("Find");
                    }
                }

                if (string.IsNullOrEmpty(ProductName))
                    throw new Exception("Fixtures audio device can not be found");

                RightSource.DeviceNumber = DeviceNumber;
                Console.WriteLine("Record device {0}",ProductName);
                RightSource.WaveFormat = new WaveFormat(44100, 1);
                RightSource.DataAvailable += new EventHandler<WaveInEventArgs>(RightSource_DataAvailable);
                RightSource.RecordingStopped += new EventHandler<StoppedEventArgs>(RightSource_RecordingStopped);
                RightSourceFile = new WaveFileWriter(RightRecordFileName, RightSource.WaveFormat);
                RightSource.StartRecording();
            }
            //catch (Exception ex)
            //{
            //    Console.WriteLine(ex);
            //    result = Result.ExceptionFail;
            //}
        }

        void RecordLeftSpeaker()
        {
            //try
            {
                LeftSource = new WaveInEvent();
                string AudioPrefix = "Logitech USB H";
                ProductName = string.Empty;

                for (int n = -1; n < WaveInEvent.DeviceCount; n++)
                {
                    var caps = WaveInEvent.GetCapabilities(n);
                    Console.WriteLine("Record device {0}: {1}", n, caps.ProductName);
                    if (caps.ProductName.Contains(AudioPrefix))
                    {
                        DeviceNumber = n;
                        ProductName = caps.ProductName;
                        Console.WriteLine("Find");
                    }
                }

                if (string.IsNullOrEmpty(ProductName))
                    throw new Exception("Fixtures audio device can not be found");

                LeftSource.DeviceNumber = DeviceNumber;
                Console.WriteLine("Record device {0}",ProductName);
                LeftSource.WaveFormat = new WaveFormat(44100, 1);
                LeftSource.DataAvailable += new EventHandler<WaveInEventArgs>(LeftSource_DataAvailable);
                LeftSource.RecordingStopped += new EventHandler<StoppedEventArgs>(LeftSource_RecordingStopped);
                LeftSourceFile = new WaveFileWriter(LeftRecordFileName, LeftSource.WaveFormat);
                LeftSource.StartRecording();
            }
            //catch (Exception ex)
            //{
            //    Console.WriteLine(ex);
            //    result = Result.ExceptionFail;
            //}
        }

        void RecordHeadSet()
        {
            //try
            {
                InternalSource = new WaveInEvent();
                string AudioPrefix = "Realtek High";
                ProductName = string.Empty;

                for (int n = -1; n < WaveInEvent.DeviceCount; n++)
                {
                    var caps = WaveInEvent.GetCapabilities(n);
                    Console.WriteLine("Record device {0}: {1}", n, caps.ProductName);
                    if (caps.ProductName.Contains(AudioPrefix))
                    {
                        DeviceNumber = n;
                        ProductName = caps.ProductName;
                        Console.WriteLine("Find");
                    }
                }

                if (string.IsNullOrEmpty(ProductName))
                    throw new Exception("Audio device can not be found");

                InternalSource.DeviceNumber = DeviceNumber;
                Console.WriteLine("Record device {0}",ProductName);
                InternalSource.WaveFormat = new WaveFormat(44100, 1);
                InternalSource.DataAvailable += new EventHandler<WaveInEventArgs>(InternalSource_DataAvailable);
                InternalSource.RecordingStopped += new EventHandler<StoppedEventArgs>(InternalSource_RecordingStopped);
                InternalSourceFile = new WaveFileWriter(InternalRecordFileName, InternalSource.WaveFormat);
                InternalSource.StartRecording();
            }
            //catch (Exception ex)
            //{
            //    Console.WriteLine(ex);
            //    result = Result.ExceptionFail;
            //}
        }

        void PlayFromSpeakerAndRecord(string WavFileName)
        {
            string AudioPrefix = "Realtek High";
            ProductName = string.Empty;
            for (int n = -1; n < WaveOut.DeviceCount; n++)
            {
                var caps = WaveOut.GetCapabilities(n);
                Console.WriteLine("Play device {0}: {1}", n, caps.ProductName);
                if (caps.ProductName.Contains(AudioPrefix))
                {
                    DeviceNumber = n;
                    ProductName = caps.ProductName;
                    Console.WriteLine("Find");
                }
            }

            if (string.IsNullOrEmpty(ProductName))
                throw new Exception("Audio device can not be found");

            using (var inputReader = new WaveFileReader(WavFileName))
            using (var outputDevice = new WaveOutEvent())
            {
                outputDevice.DeviceNumber = DeviceNumber;
                Console.WriteLine("Play device: {0}", ProductName);
                outputDevice.Init(inputReader);
                outputDevice.PlaybackStopped += new EventHandler<StoppedEventArgs>(PlayFromSpeakerStopped);
                outputDevice.Play();
                RecordLeftSpeaker();
                RecordRightSpeaker();
                while (outputDevice.PlaybackState == PlaybackState.Playing)
                {
                    Thread.Sleep(500);
                }
            }
            //catch (Exception ex)
            //{
            //    Console.WriteLine("5566"+ex);
            //    result = Result.ExceptionFail;
            //}
        }

        void PlayFromHeadSetAndRecord(string WavFileName)
        {
            string AudioPrefix = "Logitech USB H";
            ProductName = string.Empty;
            //try
            {
                for (int n = -1; n < WaveOut.DeviceCount; n++)
                {
                    var caps = WaveOut.GetCapabilities(n);
                    Console.WriteLine("Play device {0}: {1}", n, caps.ProductName);
                    if (caps.ProductName.Contains(AudioPrefix))
                    {
                        DeviceNumber = n;
                        ProductName = caps.ProductName;
                        Console.WriteLine("Find");
                    }
                }

                if (string.IsNullOrEmpty(ProductName))
                    throw new Exception("Fixtures audio device can not be found");

                using (var inputReader = new WaveFileReader(WavFileName))
                using (var outputDevice = new WaveOutEvent())
                {
                    outputDevice.DeviceNumber = DeviceNumber;
                    Console.WriteLine("Play device: {0}",ProductName);
                    outputDevice.Init(inputReader);
                    outputDevice.PlaybackStopped += new EventHandler<StoppedEventArgs>(PlayFromHeadSetStopped);
                    outputDevice.Play();
                    RecordHeadSet();
                    while (outputDevice.PlaybackState == PlaybackState.Playing)
                    {
                        Thread.Sleep(500);
                    }
                }
            }
            //catch (Exception ex)
            //{
            //    Console.WriteLine(ex);
            //    result = Result.ExceptionFail;
            //}
        }

        void PlayFromHeadSetStopped(object sender, StoppedEventArgs e)
        {
            if (InternalSource != null)
                InternalSource.StopRecording();
        }

        void PlayFromSpeakerStopped(object sender, StoppedEventArgs e)
        {
            if (LeftSource != null)
                LeftSource.StopRecording();

            if (RightSource != null)
                RightSource.StopRecording();
        }

        double CalculateRMS(string WavFileName)
        {
            using (var reader = new WaveFileReader(WavFileName))
            {
                var sp = reader.ToSampleProvider();
                var sourceSamples = (int)(reader.Length / (reader.WaveFormat.BitsPerSample / 8));
                var sampleData = new float[sourceSamples];
                var total = 0.0;
                sp.Read(sampleData, 0, sourceSamples);
                foreach (var v in sampleData)
                {
                    total += Math.Pow(v, 2);
                }

                var NormalizeRms = Math.Sqrt(total / sourceSamples);
                var Rms = NormalizeRms * Math.Pow(2, reader.WaveFormat.BitsPerSample) / 2;
                Console.WriteLine(string.Format("========{0}========", WavFileName));
                //Console.WriteLine(sourceSamples);
                //Console.WriteLine(reader.WaveFormat.BitsPerSample);
                //Console.WriteLine(reader.WaveFormat.SampleRate);
                Console.WriteLine("rms: {0}", Rms);
                return Rms;
            }

        }
    }
}
