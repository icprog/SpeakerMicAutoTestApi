using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SpeakerMicAutoTestApi
{
    public enum Result
    {
        Pass = 0,
        LeftSpeakerFail = 1,
        RightSpeakerFail = 2,
        InternalMicFail = 3
    }

    public class AudioTest
    {
        double internalthreshold { get; set; }
        double externalthreshold { get; set; }
        string wavfilename { get; set; }
        string LeftRecordFileName { get; set; }
        string RightRecordFileName { get; set; }
        string InternalRecordFileName { get; set; }
        string ProductName { get; set; }
        int DeviceNumber { get; set; }

        WaveInEvent LeftSource = null;
        WaveFileWriter LeftSourceFile = null;
        WaveInEvent RightSource = null;
        WaveFileWriter RightSourceFile = null;
        WaveInEvent InternalSource = null;
        WaveFileWriter InternalSourceFile = null;
        AutoResetEvent RecordEvent = null;
       
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

        public string WavFileName
        {
            get { return wavfilename; }
            set { wavfilename = value; }
        }

        public AudioTest()
        {
            externalthreshold = 1800.0;
            internalthreshold = 800.0;
            wavfilename = "o95.wav";
            LeftRecordFileName = "left.wav";
            RightRecordFileName = "right.wav";
            InternalRecordFileName = "headset.wav";
            DeviceNumber = 0;
            ProductName = string.Empty;
            RecordEvent = new AutoResetEvent(false);
        }

        public Result RunTest()
        {
            Task.Factory.StartNew(() =>
            {
                PlayFromSpeakerAndRecord(WavFileName);               
            });

            RecordEvent.WaitOne();

            if (CalculateRMS(LeftRecordFileName) < externalthreshold)
                return Result.LeftSpeakerFail;

            if (CalculateRMS(RightRecordFileName) < externalthreshold)
                return Result.RightSpeakerFail;

            Task.Factory.StartNew(() =>
            {
                PlayFromHeadSetAndRecord(WavFileName);
            });

            RecordEvent.WaitOne();

            if (CalculateRMS(InternalRecordFileName) < internalthreshold)
                return Result.InternalMicFail;

            return Result.Pass;
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

            RecordEvent.Set();
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

            RecordEvent.Set();
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
        }

        void RecordRightSpeaker()
        {
            try
            {
                RightSource = new WaveInEvent();
                string AudioPrefix = "(2- Logitech USB Headset";

                for (int n = -1; n < WaveInEvent.DeviceCount; n++)
                {
                    var caps = WaveInEvent.GetCapabilities(n);
                    Console.WriteLine($"{n}: {caps.ProductName}");
                    if (caps.ProductName.Contains(AudioPrefix))
                    {
                        DeviceNumber = n;
                        ProductName = caps.ProductName;
                    }
                }

                RightSource.DeviceNumber = DeviceNumber;
                Console.WriteLine(ProductName);
                RightSource.WaveFormat = new WaveFormat(44100, 1);

                RightSource.DataAvailable += new EventHandler<WaveInEventArgs>(RightSource_DataAvailable);
                RightSource.RecordingStopped += new EventHandler<StoppedEventArgs>(RightSource_RecordingStopped);
                RightSourceFile = new WaveFileWriter(RightRecordFileName, RightSource.WaveFormat);
                RightSource.StartRecording();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        void RecordLeftSpeaker()
        {
            try
            {
                LeftSource = new WaveInEvent();
                string AudioPrefix = "(Logitech USB Headset";

                for (int n = -1; n < WaveInEvent.DeviceCount; n++)
                {
                    var caps = WaveInEvent.GetCapabilities(n);
                    Console.WriteLine($"{n}: {caps.ProductName}");
                    if (caps.ProductName.Contains(AudioPrefix))
                    {
                        DeviceNumber = n;
                        ProductName = caps.ProductName;
                    }
                }

                LeftSource.DeviceNumber = DeviceNumber;
                Console.WriteLine(ProductName);
                LeftSource.WaveFormat = new WaveFormat(44100, 1);

                LeftSource.DataAvailable += new EventHandler<WaveInEventArgs>(LeftSource_DataAvailable);
                LeftSource.RecordingStopped += new EventHandler<StoppedEventArgs>(LeftSource_RecordingStopped);
                LeftSourceFile = new WaveFileWriter(LeftRecordFileName, LeftSource.WaveFormat);
                LeftSource.StartRecording();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        void RecordHeadSet()
        {
            try
            {
                InternalSource = new WaveInEvent();
                string AudioPrefix = "Realtek High";

                for (int n = -1; n < WaveInEvent.DeviceCount; n++)
                {
                    var caps = WaveInEvent.GetCapabilities(n);
                    Console.WriteLine($"{n}: {caps.ProductName}");
                    if (caps.ProductName.Contains(AudioPrefix))
                    {
                        DeviceNumber = n;
                        ProductName = caps.ProductName;
                    }
                }

                InternalSource.DeviceNumber = DeviceNumber;
                Console.WriteLine(ProductName);
                InternalSource.WaveFormat = new WaveFormat(44100, 1);

                InternalSource.DataAvailable += new EventHandler<WaveInEventArgs>(InternalSource_DataAvailable);
                InternalSource.RecordingStopped += new EventHandler<StoppedEventArgs>(InternalSource_RecordingStopped);
                InternalSourceFile = new WaveFileWriter(InternalRecordFileName, InternalSource.WaveFormat);
                InternalSource.StartRecording();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        void PlayFromSpeakerAndRecord(string WavFileName)
        {
            string AudioPrefix = "Realtek High";
            try
            {
                for (int n = -1; n < WaveOut.DeviceCount; n++)
                {
                    var caps = WaveOut.GetCapabilities(n);
                    if (caps.ProductName.Contains(AudioPrefix))
                    {
                        DeviceNumber = n;
                        ProductName = caps.ProductName;
                    }
                }

                using (var inputReader = new WaveFileReader(WavFileName))
                using (var outputDevice = new WaveOutEvent())
                {
                    outputDevice.DeviceNumber = DeviceNumber;
                    Console.WriteLine(ProductName);
                    outputDevice.Init(inputReader);
                    outputDevice.PlaybackStopped += new EventHandler<StoppedEventArgs>(PlayFromSpeakerStopped);
                    outputDevice.Play();
                    RecordLeftSpeaker();
                    RecordRightSpeaker();
                    while (outputDevice.PlaybackState == PlaybackState.Playing)
                    {
                        Thread.Sleep(500);
                        Console.WriteLine("Playing");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        void PlayFromHeadSetAndRecord(string WavFileName)
        {
            string AudioPrefix = "(Logitech USB Headset";
            try
            {
                for (int n = -1; n < WaveOut.DeviceCount; n++)
                {
                    var caps = WaveOut.GetCapabilities(n);
                    if (caps.ProductName.Contains(AudioPrefix))
                    {
                        DeviceNumber = n;
                        ProductName = caps.ProductName;
                    }
                }

                using (var inputReader = new WaveFileReader(WavFileName))
                using (var outputDevice = new WaveOutEvent())
                {
                    outputDevice.DeviceNumber = DeviceNumber;
                    Console.WriteLine(ProductName);
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
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
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
