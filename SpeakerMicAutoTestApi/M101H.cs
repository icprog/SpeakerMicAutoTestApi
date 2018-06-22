using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace SpeakerMicAutoTestApi
{
    public class AudioDeviceComparer : IComparer<string>
    {
        string pattern = @"\(\D+.*";
        public int Compare(string x, string y)
        {
            if (Regex.IsMatch(x, pattern))
                return -1;

            if (Regex.IsMatch(y, pattern))
                return 1;

            return string.Compare(x, y, true);
        }
    }

    public class M101H : M101B
    {
        protected string FanRecordFileName { get; set; }
        protected List<Guid> FanRecordDeviceList;

        public M101H()
        {
            FanRecordFileName = "Fan.wav";
            FanRecordDeviceList = GetIniValue("AUDIO", "FanRecordDeviceList").Split(',').ToList().ConvertAll(Guid.Parse);

        }

        public override Result RunTest()
        {
            try
            {
                var left = Task.Factory.StartNew(() =>
                {
                    LeftVolume = 100;
                    RightVolume = 0;
                    PlayAndRecord(WavFileName, Channel.Left);
                });

                left.Wait(7000);
                if (!left.IsCompleted)
                    throw new Exception("Play Left Speaker Timeout");
                Thread.Sleep(200);
                leftintensity = CalculateRMS(LeftRecordFileName);
                Debug.WriteLine("leftintensity {0}", leftintensity);
                Debug.WriteLine("externalthreshold {0}", externalthreshold);
                if (leftintensity < externalthreshold)
                    return Result.LeftSpeakerFail;

                var right = Task.Factory.StartNew(() =>
                {
                    LeftVolume = 0;
                    RightVolume = 100;
                    PlayAndRecord(WavFileName, Channel.Right);
                });

                right.Wait(7000);
                if (!right.IsCompleted)
                    throw new Exception("Play Right Speaker Timeout");
                Thread.Sleep(200);
                rightintensity = CalculateRMS(RightRecordFileName);
                Debug.WriteLine("rightintensity {0}", rightintensity);
                Debug.WriteLine("externalthreshold {0}", externalthreshold);
                if (rightintensity < externalthreshold)
                    return Result.RightSpeakerFail;

                var headset = Task.Factory.StartNew(() =>
                {
                    LeftVolume = 100;
                    RightVolume = 100;
                    PlayAndRecord(WavFileName, Channel.HeadSet);
                });

                headset.Wait(7000);
                if (!headset.IsCompleted)
                    throw new Exception("Play Headset Timeout");
                Thread.Sleep(200);
                internalintensity = CalculateRMS(InternalRecordFileName);
                Debug.WriteLine("internalintensity {0}", internalintensity);
                Debug.WriteLine("internalthreshold {0}", internalthreshold);
                if (internalintensity < internalthreshold)
                    return Result.InternalMicFail;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                Console.WriteLine(ex);
                exception = ex;
                return Result.ExceptionFail;
            }

            return Result.Pass;
        }

        protected override Task<string> Record(Channel Channel)
        {
            var tcs = new TaskCompletionSource<string>();
            WavSource = new WaveInEvent();
            bool IsEqual = false;
            ProductName = string.Empty;
            List<Guid> AudioDeviceList = null;
            Dictionary<int, string> di = new Dictionary<int, string>();

            switch (Channel)
            {
                case Channel.Left:
                case Channel.Right:
                    AudioDeviceList = ExternalAudioDeviceList;
                    break;
                case Channel.AudioJack:
                    AudioDeviceList = MachineAudioDeviceList;
                    break;
                case Channel.HeadSet:
                    AudioDeviceList = DigitalMicDeviceList;
                    break;
                case Channel.Fan:
                    AudioDeviceList = FanRecordDeviceList;
                    break;
            }

            for (int n = -1; n < WaveInEvent.DeviceCount; n++)
            {
                var caps = WaveInEvent.GetCapabilities(n);
                Console.WriteLine("Record device {0}: {1}", n, caps.ProductName);
                Console.WriteLine(caps.ManufacturerGuid);
                Debug.WriteLine("Record device {0}: {1}", n, caps.ProductName);
                Debug.WriteLine(caps.ManufacturerGuid);

                foreach (var v in AudioDeviceList)
                {
                    if (caps.ManufacturerGuid.Equals(v))
                    {
                        di.Add(n,caps.ProductName);
                        DeviceNumber = n;
                        ProductName = caps.ProductName;
                        Console.WriteLine("Find");
                    }
                }
            }

            switch (Channel)
            {

                case Channel.Left:
                case Channel.Right:
                    IsEqual = ExternalAudioDeviceList.Except(FanRecordDeviceList).Count() == 0;
                    if (IsEqual)
                    {
                        if (di.Count() < 2)
                            throw new Exception("External audio device not found");
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(ProductName))
                            throw new Exception("External audio device not found");
                    }

                    DeviceNumber = di.OrderBy(e => e.Value, new AudioDeviceComparer()).LastOrDefault().Key;
                    ProductName = di.OrderBy(e => e.Value, new AudioDeviceComparer()).LastOrDefault().Value;
                    break;
                case Channel.HeadSet:
                    if (string.IsNullOrEmpty(ProductName))
                        throw new Exception("Digital Mic device not found");
                    break;
                case Channel.AudioJack:
                    if (string.IsNullOrEmpty(ProductName))
                        throw new Exception("Audio Jack device not found");
                    break;
                case Channel.Fan:
                    IsEqual = ExternalAudioDeviceList.Except(FanRecordDeviceList).Count() == 0;
                    if (IsEqual)
                    {
                        if (di.Count() < 2)
                            throw new Exception("Fan record device not found");
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(ProductName))
                            throw new Exception("Fan record device not found");
                    }

                    DeviceNumber = di.OrderBy(e => e.Value, new AudioDeviceComparer()).LastOrDefault().Key;
                    ProductName = di.OrderBy(e => e.Value, new AudioDeviceComparer()).LastOrDefault().Value;
                    break;
            }

            if (string.IsNullOrEmpty(ProductName) && Channel == Channel.AudioJack)
                throw new Exception("Audio Jack device not found");
            else if (string.IsNullOrEmpty(ProductName) && (Channel == Channel.Left || Channel == Channel.Right))
                throw new Exception("External audio device not found");
            else if (string.IsNullOrEmpty(ProductName) && Channel == Channel.HeadSet)
                throw new Exception("Digital Mic device not found");

            WavSource.DeviceNumber = DeviceNumber;
            Console.WriteLine("Record device {0}", ProductName);
            WavSource.WaveFormat = new WaveFormat(44100, 1);
            WavSource.DataAvailable += (sender, e) =>
            {
                if (WavSourceFile != null)
                {
                    WavSourceFile.Write(e.Buffer, 0, e.BytesRecorded);
                    WavSourceFile.Flush();
                }
            };

            WavSource.RecordingStopped += (sender, e) =>
            {
                if (WavSource != null)
                {
                    WavSource.Dispose();
                    WavSource = null;
                }

                if (WavSourceFile != null)
                {
                    WavSourceFile.Dispose();
                    WavSourceFile = null;
                }

                Thread.Sleep(200);
                tcs.SetResult("Done");
                Console.WriteLine("Record Stopped");
            };

            switch (Channel)
            {
                case Channel.Left:
                    WavSourceFile = new WaveFileWriter(LeftRecordFileName, WavSource.WaveFormat);
                    break;
                case Channel.Right:
                    WavSourceFile = new WaveFileWriter(RightRecordFileName, WavSource.WaveFormat);
                    break;
                case Channel.HeadSet:
                    WavSourceFile = new WaveFileWriter(InternalRecordFileName, WavSource.WaveFormat);
                    break;
                case Channel.AudioJack:
                    WavSourceFile = new WaveFileWriter(AudioJackRecordFileName, WavSource.WaveFormat);
                    break;
                case Channel.Fan:
                    WavSourceFile = new WaveFileWriter(FanRecordFileName, WavSource.WaveFormat);
                    break;
            }

            WavSource.StartRecording();
            return tcs.Task;
        }

        protected override void PlayAndRecord(string WavFileName, Channel Channel)
        {
            bool IsEqual = false;
            ProductName = string.Empty;
            List<Guid> AudioDeviceList = null;
            Dictionary<int, string> di = new Dictionary<int, string>();

            switch (Channel)
            {
                case Channel.Left:
                case Channel.Right:
                case Channel.AudioJack:
                    AudioDeviceList = MachineAudioDeviceList;
                    break;
                case Channel.HeadSet:
                    AudioDeviceList = ExternalAudioDeviceList;
                    break;
                case Channel.Fan:
                    AudioDeviceList = FanRecordDeviceList;
                    break;
            }

            for (int n = -1; n < WaveOut.DeviceCount; n++)
            {
                var caps = WaveOut.GetCapabilities(n);
                Console.WriteLine("Play device {0}: {1}", n, caps.ProductName);
                Console.WriteLine(caps.ManufacturerGuid);
                Debug.WriteLine("Play device {0}: {1}", n, caps.ProductName);
                Debug.WriteLine(caps.ManufacturerGuid);
                foreach (var v in AudioDeviceList)
                {
                    if (caps.ManufacturerGuid.Equals(v))
                    {
                        di.Add(n, caps.ProductName);
                        DeviceNumber = n;
                        ProductName = caps.ProductName;
                        Console.WriteLine("Find");
                    }
                }
            }

            switch (Channel)
            {
                
                case Channel.Left:
                case Channel.Right:
                    if (string.IsNullOrEmpty(ProductName))
                        throw new Exception("Machine audio device not found");
                    break;
                case Channel.HeadSet:
                    IsEqual = ExternalAudioDeviceList.Except(FanRecordDeviceList).Count() == 0;
                    if (IsEqual)
                    {
                        if (di.Count() < 2)
                            throw new Exception("External audio device not found");
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(ProductName))
                            throw new Exception("External audio device not found");
                    }

                    DeviceNumber = di.OrderBy(e => e.Value, new AudioDeviceComparer()).FirstOrDefault().Key;
                    ProductName = di.OrderBy(e => e.Value, new AudioDeviceComparer()).FirstOrDefault().Value;
                    break;
                case Channel.AudioJack:
                    if (string.IsNullOrEmpty(ProductName))
                        throw new Exception("Audio Jack device not found");
                    break;
                case Channel.Fan:
                    IsEqual = ExternalAudioDeviceList.Except(FanRecordDeviceList).Count() == 0;
                    if (IsEqual)
                    {
                        if (di.Count() < 2)
                            throw new Exception("Fan record device not found");
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(ProductName))
                            throw new Exception("Fan record device not found");
                    }

                    DeviceNumber = di.OrderBy(e => e.Value, new AudioDeviceComparer()).LastOrDefault().Key;
                    ProductName = di.OrderBy(e => e.Value, new AudioDeviceComparer()).LastOrDefault().Value;
                    break;
            }

            using (var inputReader = new WaveFileReader(WavFileName))
            using (var outputDevice = new WaveOutEvent())
            {
                outputDevice.DeviceNumber = DeviceNumber;
                Console.WriteLine("Play device: {0}", ProductName);
                outputDevice.Init(inputReader);
                outputDevice.PlaybackStopped += (sender, e) =>
                {
                    if (WavSource != null)
                        WavSource.StopRecording();

                    Console.WriteLine("Play Stopped");
                };

                outputDevice.Play();

                switch (Channel)
                {
                    case Channel.Left:
                        Record(Channel.Left).Wait();
                        break;
                    case Channel.Right:
                        Record(Channel.Right).Wait();
                        break;
                    case Channel.AudioJack:
                        Record(Channel.AudioJack).Wait();
                        break;
                    case Channel.HeadSet:
                        Record(Channel.HeadSet).Wait();
                        break;
                    case Channel.Fan:
                        Record(Channel.Fan).Wait();
                        break;
                }

                while (outputDevice.PlaybackState == PlaybackState.Playing)
                {
                    Thread.Sleep(500);
                }
            }
        }


    }
}
