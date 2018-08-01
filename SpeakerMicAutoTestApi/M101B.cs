using IniParser;
using IniParser.Model;
using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SpeakerMicAutoTestApi
{
    public class M101B : Platform
    {
        protected string UsbAudioDeviceName { get; set; }
        protected List<Guid> MachineAudioDeviceList;
        protected List<Guid> ExternalAudioDeviceList;
        protected List<Guid> DigitalMicDeviceList;
        SetupApi SetupApi = null;

        public M101B(bool IsJsonConfig = false)
        {
            UsbAudioDeviceName = "USB Audio";

            if (IsJsonConfig)
            {
                MachineAudioDeviceList = GetConfigValue("MachineAudioDevice")
                    .Split(new string[] { "\"", "[", "]", "," }, StringSplitOptions.RemoveEmptyEntries)
                    .Where(e => !string.IsNullOrWhiteSpace(e))
                    .ToList()
                    .ConvertAll(Guid.Parse);

                ExternalAudioDeviceList = GetConfigValue("ExternalAudioDevice")
                    .Split(new string[] { "\"", "[", "]", "," }, StringSplitOptions.RemoveEmptyEntries)
                    .Where(e => !string.IsNullOrWhiteSpace(e))
                    .ToList()
                    .ConvertAll(Guid.Parse);

                DigitalMicDeviceList = GetConfigValue("DigitalMicDevice")
                    .Split(new string[] { "\"", "[", "]", "," }, StringSplitOptions.RemoveEmptyEntries)
                    .Where(e => !string.IsNullOrWhiteSpace(e))
                    .ToList()
                    .ConvertAll(Guid.Parse);
            }
            else
            {
                MachineAudioDeviceList = GetIniValue("AUDIO", "MachineAudioDevice").Split(',').ToList().ConvertAll(Guid.Parse);
                ExternalAudioDeviceList = GetIniValue("AUDIO", "ExternalAudioDevice").Split(',').ToList().ConvertAll(Guid.Parse);
                DigitalMicDeviceList = GetIniValue("AUDIO", "DigitalMicDevice").Split(',').ToList().ConvertAll(Guid.Parse);
            }            
        }

        public override Result FanTest()
        {
            throw new NotImplementedException();
        }

        public override Result AudioJackTest()
        {
            try
            {
                DeleteRecordWav();
                MicrophoneBoost = 30.0f;
                var audiojack = Task.Factory.StartNew(() =>
                {
                    LeftVolume = 100;
                    RightVolume = 100;
                    PlayAndRecord(WavFileName, Channel.AudioJack);
                });

                audiojack.Wait(AudioTimeout);
                if (!audiojack.IsCompleted)
                    throw new Exception("Play Audio Jack Timeout");

                Thread.Sleep(200);
                audiojackintensity = CalculateRMS(AudioJackRecordFileName);
                if (audiojackintensity < audiojackthreshold)
                    return Result.AudioJackFail;
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex);
                Console.WriteLine(ex);
                exception = ex;
                return Result.ExceptionFail;
            }
            finally
            {
                DeleteRecordWav();
                MicrophoneBoost = 0.0f;
            }

            return Result.Pass;
        }

        public override Result RunTest()
        {
            try
            {
                DeleteRecordWav();
                MicrophoneBoost = 30.0f;
                var left = Task.Factory.StartNew(() =>
                {
                    LeftVolume = 100;
                    RightVolume = 0;
                    PlayAndRecord(WavFileName, Channel.Left);
                });

                left.Wait(AudioTimeout);
                if (!left.IsCompleted)
                    throw new Exception("Play Left Speaker Timeout");
                Thread.Sleep(200);
                leftintensity = CalculateRMS(LeftRecordFileName);
                if (leftintensity < externalthreshold)
                    return Result.LeftSpeakerFail;

                var right = Task.Factory.StartNew(() =>
                {
                    LeftVolume = 0;
                    RightVolume = 100;
                    PlayAndRecord(WavFileName, Channel.Right);
                });

                right.Wait(AudioTimeout);
                if (!right.IsCompleted)
                    throw new Exception("Play Right Speaker Timeout");
                Thread.Sleep(200);
                rightintensity = CalculateRMS(RightRecordFileName);
                if (rightintensity < externalthreshold)
                    return Result.RightSpeakerFail;

                var headset = Task.Factory.StartNew(() =>
                {
                    LeftVolume = 100;
                    RightVolume = 100;
                    PlayAndRecord(WavFileName, Channel.HeadSet);
                });

                headset.Wait(AudioTimeout);
                if (!headset.IsCompleted)
                    throw new Exception("Play Headset Timeout");
                Thread.Sleep(200);
                internalintensity = CalculateRMS(InternalRecordFileName);
                internalleftintensity = internalintensity;
                internalrightintensity = internalintensity;
                if (internalintensity < internalthreshold)
                    return Result.InternalMicFail;
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex);
                Console.WriteLine(ex);
                exception = ex;
                return Result.ExceptionFail;
            }
            finally
            {
                DeleteRecordWav();
                MicrophoneBoost = 0.0f;
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
            }

            SetupApi.di.Clear();
            for (int n = -1; n < WaveInEvent.DeviceCount; n++)
            {
                var caps = WaveInEvent.GetCapabilities(n);
                Console.WriteLine("Record device {0}: {1}", n, caps.ProductName);
                Console.WriteLine(caps.ManufacturerGuid);
                Trace.WriteLine(caps.ManufacturerGuid);

                foreach (var v in AudioDeviceList)
                {
                    if (caps.ManufacturerGuid.Equals(v))
                    {
                        DeviceNumber = n;
                        ProductName = caps.ProductName;
                        SetupApi.GetLocationInformation(DeviceNumber, ProductName);
                        Console.WriteLine("Find");
                    }
                }
            }

            switch (Channel)
            {
                case Channel.Left:
                case Channel.Right:

                    if (string.IsNullOrEmpty(ProductName) || ProductName.Contains(UsbAudioDeviceName))
                        throw new Exception("External audio device not found");

                    DeviceNumber = SetupApi.di.OrderBy(e => e.Value).FirstOrDefault().Key;
                    //ProductName = SetupApi.di.OrderBy(e => e.Value).FirstOrDefault().Value;
                    break;
                case Channel.HeadSet:
                    if (string.IsNullOrEmpty(ProductName))
                        throw new Exception("Digital Mic device not found");
                    break;
                case Channel.AudioJack:
                    if (string.IsNullOrEmpty(ProductName))
                        throw new Exception("Audio Jack device not found");
                    break;
            }

            WavSource.DeviceNumber = DeviceNumber;
            Console.WriteLine("Record device ###### {0} ######", WaveInEvent.GetCapabilities(DeviceNumber).ProductName);
            foreach (var item in SetupApi.di.OrderBy(e => e.Value))
            {
                Console.WriteLine(item.Key + "   " + item.Value);
            }
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
            SetupApi = new SetupApi();

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
            }

            for (int n = -1; n < WaveOut.DeviceCount; n++)
            {
                var caps = WaveOut.GetCapabilities(n);
                Console.WriteLine("Play device {0}: {1}", n, caps.ProductName);
                Console.WriteLine(caps.ManufacturerGuid);
                Trace.WriteLine(caps.ManufacturerGuid);
                foreach (var v in AudioDeviceList)
                {
                    if (caps.ManufacturerGuid.Equals(v))
                    {
                        DeviceNumber = n;
                        ProductName = caps.ProductName;
                        SetupApi.GetLocationInformation(DeviceNumber, ProductName);
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

                    if (string.IsNullOrEmpty(ProductName) || ProductName.Contains(UsbAudioDeviceName))
                        throw new Exception("External audio device not found");

                    DeviceNumber = SetupApi.di.OrderBy(e => e.Value).FirstOrDefault().Key;
                    //ProductName = SetupApi.di.OrderBy(e => e.Value).FirstOrDefault().Value;
                    break;
                case Channel.AudioJack:
                    if (string.IsNullOrEmpty(ProductName))
                        throw new Exception("Audio Jack device not found");
                    break;
            }

            using (var inputReader = new WaveFileReader(WavFileName))
            using (var outputDevice = new WaveOutEvent())
            {
                outputDevice.DeviceNumber = DeviceNumber;
                Console.WriteLine("Play device: ###### {0} ######", WaveOut.GetCapabilities(DeviceNumber).ProductName);
                foreach (var item in SetupApi.di.OrderBy(e => e.Value))
                {
                    Console.WriteLine(item.Key + "   " + item.Value);
                }
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

        protected override double CalculateRMS(string WavFileName)
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
                Console.WriteLine("rms: {0}", Rms);
                return Rms;
            }
        }
    }

    class M101BPlatform : PlatformFactory
    {
        public Platform Create()
        {
            return new M101B();
        }
    }
}
