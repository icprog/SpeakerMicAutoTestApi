using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces;
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
            externalthreshold = 18000.0;
            internalthreshold = 18000.0;
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
                    PlayFromSpeakerAndRecord(WavFileName, 0);
                });

                left.Wait();
                leftintensity = CalculateRMS(LeftRecordFileName);
                if (leftintensity < externalthreshold)
                    result = Result.LeftSpeakerFail;

                var right = Task.Factory.StartNew(() =>
                {
                    LeftVolume = 0;
                    RightVolume = 100;
                    PlayFromSpeakerAndRecord(WavFileName, 1);
                });

                right.Wait();
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

        Task<string> RecordRightSpeaker()
        {
            var tcs = new TaskCompletionSource<string>();
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
            Console.WriteLine("Record device {0}", ProductName);
            RightSource.WaveFormat = new WaveFormat(44100, 1);
            RightSource.DataAvailable += (sender, e) =>
            {
                if (RightSourceFile != null)
                {
                    RightSourceFile.Write(e.Buffer, 0, e.BytesRecorded);
                    RightSourceFile.Flush();
                }
            };

            RightSource.RecordingStopped += (sender, e) =>
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
                tcs.SetResult("Done");
                Console.WriteLine("Right record Stopped");
            };

            RightSourceFile = new WaveFileWriter(RightRecordFileName, RightSource.WaveFormat);
            RightSource.StartRecording();
            return tcs.Task;
        }

        Task<string> RecordLeftSpeaker()
        {
            var tcs = new TaskCompletionSource<string>();
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
            Console.WriteLine("Record device {0}", ProductName);
            LeftSource.WaveFormat = new WaveFormat(44100, 1);
            LeftSource.DataAvailable += (sender, e) =>
            {
                if (LeftSourceFile != null)
                {
                    LeftSourceFile.Write(e.Buffer, 0, e.BytesRecorded);
                    LeftSourceFile.Flush();
                }
            };

            LeftSource.RecordingStopped += (sender, e) =>
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
                tcs.SetResult("Done");
                Console.WriteLine("Left record Stopped");
            };

            LeftSourceFile = new WaveFileWriter(LeftRecordFileName, LeftSource.WaveFormat);
            LeftSource.StartRecording();
            return tcs.Task;
        }

        Task<string> RecordHeadSet()
        {
            var tcs = new TaskCompletionSource<string>();
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
            Console.WriteLine("Record device: {0}", ProductName);
            InternalSource.WaveFormat = new WaveFormat(44100, 1);
            InternalSource.DataAvailable += (sender, e) =>
            {
                if (InternalSourceFile != null)
                {
                    InternalSourceFile.Write(e.Buffer, 0, e.BytesRecorded);
                    InternalSourceFile.Flush();
                }
            };

            InternalSource.RecordingStopped += (sender, e) =>
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
                tcs.SetResult("Done");
                Console.WriteLine("Internal record Stopped");
            };

            InternalSourceFile = new WaveFileWriter(InternalRecordFileName, InternalSource.WaveFormat);
            InternalSource.StartRecording();
            return tcs.Task;
        }

        void PlayFromSpeakerAndRecord(string WavFileName, int channel)
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
                outputDevice.PlaybackStopped += (sender, e) =>
                {
                    if (LeftSource != null)
                        LeftSource.StopRecording();

                    if (RightSource != null)
                        RightSource.StopRecording();

                    Console.WriteLine("Speaker Play Stopped");
                };

                outputDevice.Play();
                if (channel == 0)
                    RecordLeftSpeaker().Wait();
                else
                    RecordRightSpeaker().Wait();

                while (outputDevice.PlaybackState == PlaybackState.Playing)
                {
                    Thread.Sleep(500);
                }
            }
        }

        void PlayFromHeadSetAndRecord(string WavFileName)
        {
            string AudioPrefix = "Logitech USB H";
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
                throw new Exception("Fixtures audio device can not be found");

            using (var inputReader = new WaveFileReader(WavFileName))
            using (var outputDevice = new WaveOutEvent())
            {
                outputDevice.DeviceNumber = DeviceNumber;
                Console.WriteLine("Play device: {0}", ProductName);
                outputDevice.Init(inputReader);
                outputDevice.PlaybackStopped += (sender, e) =>
                {
                    if (InternalSource != null)
                        InternalSource.StopRecording();

                    Console.WriteLine("Headset Play Stopped");
                };

                outputDevice.Play();
                RecordHeadSet().Wait();
                while (outputDevice.PlaybackState == PlaybackState.Playing)
                {
                    Thread.Sleep(500);
                }
            }
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
