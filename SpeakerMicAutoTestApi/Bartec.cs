using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SpeakerMicAutoTestApi
{
    public class Bartec : Platform
    {
        List<Guid> MachineAudioDeviceList;
        List<Guid> ExternalAudioDeviceList;

        public Bartec(bool IsJsonConfig = false)
        {
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
            }
            else
            {
                MachineAudioDeviceList = GetIniValue("AUDIO", "MachineAudioDevice").Split(',').ToList().ConvertAll(Guid.Parse);
                ExternalAudioDeviceList = GetIniValue("AUDIO", "ExternalAudioDevice").Split(',').ToList().ConvertAll(Guid.Parse);
            }
        }

        public override Result FanTest()
        {
            throw new NotImplementedException();
        }

        public override Result AudioJackTest()
        {
            throw new NotImplementedException();
        }

        protected override Task<string> Record(Channel Channel)
        {
            var tcs = new TaskCompletionSource<string>();
            WavSource = new WaveInEvent();
            ProductName = string.Empty;
            List<Guid> AudioDeviceList = null;

            if (Channel == Channel.HeadSet)
            {
                AudioDeviceList = MachineAudioDeviceList;
            }
            else if (Channel == Channel.Left || Channel == Channel.Right)
            {
                AudioDeviceList = ExternalAudioDeviceList;
            }

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
                        Console.WriteLine("Find");
                    }
                }
            }

            if (string.IsNullOrEmpty(ProductName) && Channel == Channel.HeadSet)
                throw new Exception("Machine audio device not found");
            else if (string.IsNullOrEmpty(ProductName) && (Channel == Channel.Left || Channel == Channel.Right))
                throw new Exception("External audio device not found");

            WavSource.DeviceNumber = DeviceNumber;
            Console.WriteLine("Record device {0}", ProductName);
            WavSource.WaveFormat = new WaveFormat(44100, 2);
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

            WavSource.StartRecording();
            return tcs.Task;
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
                SplitTwoChannel(InternalRecordFileName);
                Thread.Sleep(200);
                internalleftintensity = CalculateRMS(LeftChannelFileName);
                if (internalleftintensity < internalthreshold)
                    return Result.InternalLeftMicFail;

                internalrightintensity = CalculateRMS(RightChannelFileName);
                if (internalrightintensity < internalthreshold)
                    return Result.InternalRightMicFail;
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
                LeftVolume = 100;
                RightVolume = 100;
            }

            return Result.Pass;
        }

        protected override void PlayAndRecord(string WavFileName, Channel Channel)
        {
            ProductName = string.Empty;
            List<Guid> AudioDeviceList = null;

            if (Channel == Channel.Left || Channel == Channel.Right)
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
                Trace.WriteLine(caps.ManufacturerGuid);
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
                else if (Channel == Channel.HeadSet)
                    Record(Channel.HeadSet).Wait();

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

        protected void SplitTwoChannel(string FileName)
        {
            int bytesRead;
            var reader = new WaveFileReader(FileName);
            var buffer = new byte[2 * reader.WaveFormat.SampleRate * reader.WaveFormat.Channels];
            var writers = new WaveFileWriter[reader.WaveFormat.Channels];
            for (int n = 0; n < writers.Length; n++)
            {
                var format = new WaveFormat(reader.WaveFormat.SampleRate, 16, 1);
                writers[n] = new WaveFileWriter(GetFullPath(String.Format("channel{0}.wav", n + 1)), format);
            }

            while ((bytesRead = reader.Read(buffer, 0, buffer.Length)) > 0)
            {
                int offset = 0;
                while (offset < bytesRead)
                {
                    for (int n = 0; n < writers.Length; n++)
                    {
                        writers[n].Write(buffer, offset, 2);
                        offset += 2;
                    }
                }
            }

            for (int n = 0; n < writers.Length; n++)
            {
                writers[n].Dispose();
            }
            reader.Dispose();
        }
    }

    class BartecPlatform : PlatformFactory
    {
        public Platform Create()
        {
            return new Bartec();
        }
    }
}
