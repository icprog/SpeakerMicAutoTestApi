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
    public class M101B : Platform
    {
        string AudioJackRecordFileName { get; set; }

        string MachineAudioDeviceName = "Realtek High";
        string MachineRecordDeviceName = "USB Aud";
        string ExternalAudioDeviceName = "Logitech USB H";
        string AudioJackAudioDeviceName = "Mic in";
        
        WaveInEvent AudioJackSource = null;
        WaveFileWriter AudioJackSourceFile = null;

        public M101B()
        {
            AudioJackRecordFileName = "audiojack.wav";
            externalthreshold = 20.0;
            internalthreshold = 20.0;
            audiojackthreshold = 20.0;
        }

        public override Result AudioJackTest()
        {
            try
            {
                var audiojack = Task.Factory.StartNew(() =>
                {
                    LeftVolume = 100;
                    RightVolume = 100;
                    PlayFromAudioJackAndRecord(WavFileName);
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
                    PlayFromSpeakerAndRecord(WavFileName, 0);
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
                    PlayFromSpeakerAndRecord(WavFileName, 1);
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
                    PlayFromHeadSetAndRecord(WavFileName);
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

        protected override Task<string> RecordRightSpeaker()
        {
            var tcs = new TaskCompletionSource<string>();
            RightSource = new WaveInEvent();
            ProductName = string.Empty;

            for (int n = -1; n < WaveInEvent.DeviceCount; n++)
            {
                var caps = WaveInEvent.GetCapabilities(n);
                Console.WriteLine("Record device {0}: {1}", n, caps.ProductName);
                if (caps.ProductName.Contains(ExternalAudioDeviceName))
                {
                    DeviceNumber = n;
                    ProductName = caps.ProductName;
                    Console.WriteLine("Find");
                }
            }

            if (string.IsNullOrEmpty(ProductName))
                throw new Exception("External audio device can not be found");

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

        protected override Task<string> RecordLeftSpeaker()
        {
            var tcs = new TaskCompletionSource<string>();
            LeftSource = new WaveInEvent();
            ProductName = string.Empty;

            for (int n = -1; n < WaveInEvent.DeviceCount; n++)
            {
                var caps = WaveInEvent.GetCapabilities(n);
                Console.WriteLine("Record device {0}: {1}", n, caps.ProductName);
                if (caps.ProductName.Contains(ExternalAudioDeviceName))
                {
                    DeviceNumber = n;
                    ProductName = caps.ProductName;
                    Console.WriteLine("Find");
                }
            }

            if (string.IsNullOrEmpty(ProductName))
                throw new Exception("External audio device can not be found");

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

        protected override Task<string> RecordHeadSet()
        {
            var tcs = new TaskCompletionSource<string>();
            InternalSource = new WaveInEvent();
            ProductName = string.Empty;

            for (int n = -1; n < WaveInEvent.DeviceCount; n++)
            {
                var caps = WaveInEvent.GetCapabilities(n);
                Console.WriteLine("Record device {0}: {1}", n, caps.ProductName);
                if (caps.ProductName.Contains(MachineRecordDeviceName))
                {
                    DeviceNumber = n;
                    ProductName = caps.ProductName;
                    Console.WriteLine("Find");
                }
            }

            if (string.IsNullOrEmpty(ProductName))
                throw new Exception("Machine Audio device can not be found");

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

        protected Task<string> RecordAudioJack()
        {
            var tcs = new TaskCompletionSource<string>();
            AudioJackSource = new WaveInEvent();
            ProductName = string.Empty;

            for (int n = -1; n < WaveInEvent.DeviceCount; n++)
            {
                var caps = WaveInEvent.GetCapabilities(n);
                Console.WriteLine("Record device {0}: {1}", n, caps.ProductName);
                if (caps.ProductName.Contains(AudioJackAudioDeviceName))
                {
                    DeviceNumber = n;
                    ProductName = caps.ProductName;
                    Console.WriteLine("Find");
                }
            }

            if (string.IsNullOrEmpty(ProductName))
                throw new Exception("Audio jack device can not be found");

            AudioJackSource.DeviceNumber = DeviceNumber;
            Console.WriteLine("Record device {0}", ProductName);
            AudioJackSource.WaveFormat = new WaveFormat(44100, 1);
            AudioJackSource.DataAvailable += (sender, e) =>
            {
                if (AudioJackSourceFile != null)
                {
                    AudioJackSourceFile.Write(e.Buffer, 0, e.BytesRecorded);
                    AudioJackSourceFile.Flush();
                }
            };

            AudioJackSource.RecordingStopped += (sender, e) =>
            {
                if (AudioJackSource != null)
                {
                    AudioJackSource.Dispose();
                    AudioJackSource = null;
                }

                if (AudioJackSourceFile != null)
                {
                    AudioJackSourceFile.Dispose();
                    AudioJackSourceFile = null;
                }

                Thread.Sleep(200);
                tcs.SetResult("Done");
                Console.WriteLine("Audio jack record Stopped");
            };

            AudioJackSourceFile = new WaveFileWriter(AudioJackRecordFileName, AudioJackSource.WaveFormat);
            AudioJackSource.StartRecording();
            return tcs.Task;
        }


        protected void PlayFromAudioJackAndRecord(string WavFileName)
        {
            ProductName = string.Empty;
            for (int n = -1; n < WaveOut.DeviceCount; n++)
            {
                var caps = WaveOut.GetCapabilities(n);
                Console.WriteLine("Play device {0}: {1}", n, caps.ProductName);
                if (caps.ProductName.Contains(MachineAudioDeviceName))
                {
                    DeviceNumber = n;
                    ProductName = caps.ProductName;
                    Console.WriteLine("Find");
                }
            }

            if (string.IsNullOrEmpty(ProductName))
                throw new Exception("Audio jack device can not be found");

            using (var inputReader = new WaveFileReader(WavFileName))
            using (var outputDevice = new WaveOutEvent())
            {
                outputDevice.DeviceNumber = DeviceNumber;
                Console.WriteLine("Play device: {0}", ProductName);
                outputDevice.Init(inputReader);
                outputDevice.PlaybackStopped += (sender, e) =>
                {
                    if (AudioJackSource != null)
                        AudioJackSource.StopRecording();

                    Console.WriteLine("Audio jack Play Stopped");
                };

                outputDevice.Play();
                RecordAudioJack().Wait();
                while (outputDevice.PlaybackState == PlaybackState.Playing)
                {
                    Thread.Sleep(500);
                }
            }
        }

        protected override void PlayFromSpeakerAndRecord(string WavFileName, int channel)
        {
            ProductName = string.Empty;
            for (int n = -1; n < WaveOut.DeviceCount; n++)
            {
                var caps = WaveOut.GetCapabilities(n);
                Console.WriteLine("Play device {0}: {1}", n, caps.ProductName);
                if (caps.ProductName.Contains(MachineAudioDeviceName))
                {
                    DeviceNumber = n;
                    ProductName = caps.ProductName;
                    Console.WriteLine("Find");
                }
            }

            if (string.IsNullOrEmpty(ProductName))
                throw new Exception("Machine Audio device can not be found");

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

        protected override void PlayFromHeadSetAndRecord(string WavFileName)
        {
            ProductName = string.Empty;
            for (int n = -1; n < WaveOut.DeviceCount; n++)
            {
                var caps = WaveOut.GetCapabilities(n);
                Console.WriteLine("Play device {0}: {1}", n, caps.ProductName);
                if (caps.ProductName.Contains(ExternalAudioDeviceName))
                {
                    DeviceNumber = n;
                    ProductName = caps.ProductName;
                    Console.WriteLine("Find");
                }
            }

            if (string.IsNullOrEmpty(ProductName))
                throw new Exception("External audio device can not be found");

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
