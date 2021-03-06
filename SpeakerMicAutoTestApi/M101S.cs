﻿using NAudio.Wave;
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
    public class M101S : M101H
    {
        protected string RealtekDigitalMicrophone { get; set; }

        public M101S(bool IsJsonConfig = false)
        {
            RealtekDigitalMicrophone = "Digital Mic";
            di = new Dictionary<int, string>();
        }

        protected override void PlayAndRecord(string WavFileName, Channel Channel)
        {
            bool IsEqual = false;
            ProductName = string.Empty;
            List<Guid> AudioDeviceList = null;
            SetupApi = new SetupApi();
            di.Clear();

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
                    IsEqual = ExternalAudioDeviceList.Except(FanRecordDeviceList).Count() == 0;
                    if (IsEqual)
                    {
                        if (SetupApi.di.Count() < 2 || ProductName.Contains(UsbAudioDeviceName))
                            throw new Exception("External audio device not found");
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(ProductName) || ProductName.Contains(UsbAudioDeviceName))
                            throw new Exception("External audio device not found");
                    }

                    DeviceNumber = SetupApi.di.OrderBy(e => e.Value).FirstOrDefault().Key;
                    //ProductName = SetupApi.di.OrderBy(e => e.Value).FirstOrDefault().Value;
                    break;
                case Channel.AudioJack:
                    if (string.IsNullOrEmpty(ProductName))
                        throw new Exception("Audio Jack device not found");
                    break;
                case Channel.Fan:
                    IsEqual = ExternalAudioDeviceList.Except(FanRecordDeviceList).Count() == 0;
                    if (IsEqual)
                    {
                        if (SetupApi.di.Count() < 2 || ProductName.Contains(UsbAudioDeviceName))
                            throw new Exception("Fan record device not found");
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(ProductName) || ProductName.Contains(UsbAudioDeviceName))
                            throw new Exception("Fan record device not found");
                    }

                    DeviceNumber = SetupApi.di.OrderBy(e => e.Value).LastOrDefault().Key;
                    //ProductName = SetupApi.di.OrderBy(e => e.Value).LastOrDefault().Value;
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
                case Channel.Fan:
                    AudioDeviceList = FanRecordDeviceList;
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

                        if (!di.ContainsKey(DeviceNumber))
                            di.Add(DeviceNumber, ProductName);

                        SetupApi.GetLocationInformation(DeviceNumber, ProductName);
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
                        if (SetupApi.di.Count() < 2 || ProductName.Contains(UsbAudioDeviceName))
                            throw new Exception("External audio device not found");
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(ProductName) || ProductName.Contains(UsbAudioDeviceName))
                            throw new Exception("External audio device not found");
                    }

                    DeviceNumber = SetupApi.di.OrderBy(e => e.Value).FirstOrDefault().Key;
                    //ProductName = SetupApi.di.OrderBy(e => e.Value).FirstOrDefault().Value;
                    break;
                case Channel.HeadSet:
                    if (string.IsNullOrEmpty(ProductName))
                        throw new Exception("Digital Mic device not found");
                    break;
                case Channel.AudioJack:
                    ProductName = string.Empty;
                    Regex regex = new Regex(RealtekMicrophone);
                    foreach (var v in di)
                    {
                        if (regex.IsMatch(v.Value))
                        {
                            DeviceNumber = v.Key;
                            ProductName = v.Value;
                        }
                    }

                    if (string.IsNullOrEmpty(ProductName))
                        throw new Exception("Audio Jack device not found");
                    break;
                case Channel.Fan:
                    IsEqual = ExternalAudioDeviceList.Except(FanRecordDeviceList).Count() == 0;
                    if (IsEqual)
                    {
                        if (SetupApi.di.Count() < 2)
                            throw new Exception("Fan record device not found");
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(ProductName))
                            throw new Exception("Fan record device not found");
                    }

                    DeviceNumber = SetupApi.di.OrderBy(e => e.Value).LastOrDefault().Key;
                    //ProductName = SetupApi.di.OrderBy(e => e.Value).LastOrDefault().Value;
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
    }
}
