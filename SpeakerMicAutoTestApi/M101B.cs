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
        protected string AudioJackRecordFileName { get; set; }
        protected List<Guid> MachineAudioDeviceList;
        protected List<Guid> ExternalAudioDeviceList;
        protected List<Guid> DigitalMicDeviceList;

        public M101B()
        {
            AudioJackRecordFileName = "audiojack.wav";
            MachineAudioDeviceList = GetIniValue("AUDIO", "MachineAudioDevice").Split(',').ToList().ConvertAll(Guid.Parse);
            ExternalAudioDeviceList = GetIniValue("AUDIO", "ExternalAudioDevice").Split(',').ToList().ConvertAll(Guid.Parse);
            DigitalMicDeviceList = GetIniValue("AUDIO", "DigitalMicDevice").Split(',').ToList().ConvertAll(Guid.Parse);
        }

        protected override void PlayAndRecord(string WavFileName, Channel Channel)
        {
            ProductName = string.Empty;
            List<Guid> AudioDeviceList = null;

            if (Channel == Channel.Left || Channel == Channel.Right || Channel == Channel.AudioJack)
            {
                AudioDeviceList = MachineAudioDeviceList;
            }
            else if (Channel == Channel.HeadSet)
            {
                AudioDeviceList = ExternalAudioDeviceList;
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
                        DeviceNumber = n;
                        ProductName = caps.ProductName;
                        Console.WriteLine("Find");
                    }
                }
            }

            if (string.IsNullOrEmpty(ProductName) && (Channel == Channel.Left || Channel == Channel.Right))
                throw new Exception("Machine audio device not found");
            else if (string.IsNullOrEmpty(ProductName) && Channel == Channel.HeadSet)
                throw new Exception("External audio device not found");
            else if (string.IsNullOrEmpty(ProductName) && Channel == Channel.AudioJack)
                throw new Exception("Audio Jack device not found");

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
                if (Channel == Channel.Left)
                    Record(Channel.Left).Wait();
                else if (Channel == Channel.Right)
                    Record(Channel.Right).Wait();
                else if (Channel == Channel.AudioJack)
                    Record(Channel.AudioJack).Wait();
                else if (Channel == Channel.HeadSet)
                    Record(Channel.HeadSet).Wait();

                while (outputDevice.PlaybackState == PlaybackState.Playing)
                {
                    Thread.Sleep(500);
                }
            }
        }

        public override Result AudioJackTest()
        {
            try
            {
                var audiojack = Task.Factory.StartNew(() =>
                {
                    LeftVolume = 100;
                    RightVolume = 100;
                    PlayAndRecord(WavFileName, Channel.AudioJack);
                });

                audiojack.Wait(7000);
                if (!audiojack.IsCompleted)
                    throw new Exception("Play Audio Jack Timeout");

                Thread.Sleep(200);
                audiojackintensity = CalculateRMS(AudioJackRecordFileName);
                Debug.WriteLine("audiojackintensity {0}", audiojackintensity);
                Debug.WriteLine("audiojackthreshold {0}", audiojackthreshold);
                if (audiojackintensity < audiojackthreshold)
                    return Result.AudioJackFail;
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
            ProductName = string.Empty;
            List<Guid> AudioDeviceList = null;

            if (Channel == Channel.AudioJack)
            {
                AudioDeviceList = MachineAudioDeviceList;
            }
            else if (Channel == Channel.Left || Channel == Channel.Right)
            {
                AudioDeviceList = ExternalAudioDeviceList;
            }
            else if (Channel == Channel.HeadSet)
            {
                AudioDeviceList = DigitalMicDeviceList;
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
                        DeviceNumber = n;
                        ProductName = caps.ProductName;
                        Console.WriteLine("Find");
                    }
                }
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

            if (Channel == Channel.Left)
                WavSourceFile = new WaveFileWriter(LeftRecordFileName, WavSource.WaveFormat);
            else if (Channel == Channel.Right)
                WavSourceFile = new WaveFileWriter(RightRecordFileName, WavSource.WaveFormat);
            else if (Channel == Channel.HeadSet)
                WavSourceFile = new WaveFileWriter(InternalRecordFileName, WavSource.WaveFormat);
            else if (Channel == Channel.AudioJack)
                WavSourceFile = new WaveFileWriter(AudioJackRecordFileName, WavSource.WaveFormat);

            WavSource.StartRecording();
            return tcs.Task;
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
