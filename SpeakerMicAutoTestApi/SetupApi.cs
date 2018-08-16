using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace SpeakerMicAutoTestApi
{
    class SetupApi
    {
        static readonly IntPtr INVALID_HANDLE_VALUE = new IntPtr(-1);

        /// <summary>
        /// Device registry property codes
        /// </summary>
        public enum SPDRP : int
        {
            /// <summary>
            /// DeviceDesc (R/W)
            /// </summary>
            SPDRP_DEVICEDESC = 0x00000000,

            /// <summary>
            /// HardwareID (R/W)
            /// </summary>
            SPDRP_HARDWAREID = 0x00000001,

            /// <summary>
            /// CompatibleIDs (R/W)
            /// </summary>
            SPDRP_COMPATIBLEIDS = 0x00000002,

            /// <summary>
            /// unused
            /// </summary>
            SPDRP_UNUSED0 = 0x00000003,

            /// <summary>
            /// Service (R/W)
            /// </summary>
            SPDRP_SERVICE = 0x00000004,

            /// <summary>
            /// unused
            /// </summary>
            SPDRP_UNUSED1 = 0x00000005,

            /// <summary>
            /// unused
            /// </summary>
            SPDRP_UNUSED2 = 0x00000006,

            /// <summary>
            /// Class (R--tied to ClassGUID)
            /// </summary>
            SPDRP_CLASS = 0x00000007,

            /// <summary>
            /// ClassGUID (R/W)
            /// </summary>
            SPDRP_CLASSGUID = 0x00000008,

            /// <summary>
            /// Driver (R/W)
            /// </summary>
            SPDRP_DRIVER = 0x00000009,

            /// <summary>
            /// ConfigFlags (R/W)
            /// </summary>
            SPDRP_CONFIGFLAGS = 0x0000000A,

            /// <summary>
            /// Mfg (R/W)
            /// </summary>
            SPDRP_MFG = 0x0000000B,

            /// <summary>
            /// FriendlyName (R/W)
            /// </summary>
            SPDRP_FRIENDLYNAME = 0x0000000C,

            /// <summary>
            /// LocationInformation (R/W)
            /// </summary>
            SPDRP_LOCATION_INFORMATION = 0x0000000D,

            /// <summary>
            /// PhysicalDeviceObjectName (R)
            /// </summary>
            SPDRP_PHYSICAL_DEVICE_OBJECT_NAME = 0x0000000E,

            /// <summary>
            /// Capabilities (R)
            /// </summary>
            SPDRP_CAPABILITIES = 0x0000000F,

            /// <summary>
            /// UiNumber (R)
            /// </summary>
            SPDRP_UI_NUMBER = 0x00000010,

            /// <summary>
            /// UpperFilters (R/W)
            /// </summary>
            SPDRP_UPPERFILTERS = 0x00000011,

            /// <summary>
            /// LowerFilters (R/W)
            /// </summary>
            SPDRP_LOWERFILTERS = 0x00000012,

            /// <summary>
            /// BusTypeGUID (R)
            /// </summary>
            SPDRP_BUSTYPEGUID = 0x00000013,

            /// <summary>
            /// LegacyBusType (R)
            /// </summary>
            SPDRP_LEGACYBUSTYPE = 0x00000014,

            /// <summary>
            /// BusNumber (R)
            /// </summary>
            SPDRP_BUSNUMBER = 0x00000015,

            /// <summary>
            /// Enumerator Name (R)
            /// </summary>
            SPDRP_ENUMERATOR_NAME = 0x00000016,

            /// <summary>
            /// Security (R/W, binary form)
            /// </summary>
            SPDRP_SECURITY = 0x00000017,

            /// <summary>
            /// Security (W, SDS form)
            /// </summary>
            SPDRP_SECURITY_SDS = 0x00000018,

            /// <summary>
            /// Device Type (R/W)
            /// </summary>
            SPDRP_DEVTYPE = 0x00000019,

            /// <summary>
            /// Device is exclusive-access (R/W)
            /// </summary>
            SPDRP_EXCLUSIVE = 0x0000001A,

            /// <summary>
            /// Device Characteristics (R/W)
            /// </summary>
            SPDRP_CHARACTERISTICS = 0x0000001B,

            /// <summary>
            /// Device Address (R)
            /// </summary>
            SPDRP_ADDRESS = 0x0000001C,

            /// <summary>
            /// UiNumberDescFormat (R/W)
            /// </summary>
            SPDRP_UI_NUMBER_DESC_FORMAT = 0X0000001D,

            /// <summary>
            /// Device Power Data (R)
            /// </summary>
            SPDRP_DEVICE_POWER_DATA = 0x0000001E,

            /// <summary>
            /// Removal Policy (R)
            /// </summary>
            SPDRP_REMOVAL_POLICY = 0x0000001F,

            /// <summary>
            /// Hardware Removal Policy (R)
            /// </summary>
            SPDRP_REMOVAL_POLICY_HW_DEFAULT = 0x00000020,

            /// <summary>
            /// Removal Policy Override (RW)
            /// </summary>
            SPDRP_REMOVAL_POLICY_OVERRIDE = 0x00000021,

            /// <summary>
            /// Device Install State (R)
            /// </summary>
            SPDRP_INSTALL_STATE = 0x00000022,

            /// <summary>
            /// Device Location Paths (R)
            /// </summary>
            SPDRP_LOCATION_PATHS = 0x00000023,
        }

        [Flags]
        enum DiGetClassFlags : int
        {
            DIGCF_DEFAULT = 0x00000001, // only valid with DIGCF_DEVICEINTERFACE
            DIGCF_PRESENT = 0x00000002,
            DIGCF_ALLCLASSES = 0x00000004,
            DIGCF_PROFILE = 0x00000008,
            DIGCF_DEVICEINTERFACE = 0x00000010,
        }

        [Flags]
        enum DEVPROPTYPE : ulong
        {
            DEVPROP_TYPEMOD_ARRAY = 0x00001000,
            DEVPROP_TYPEMOD_LIST = 0x00002000,

            DEVPROP_TYPE_EMPTY = 0x00000000,  // nothing, no property data
            DEVPROP_TYPE_NULL = 0x00000001,  // null property data
            DEVPROP_TYPE_SBYTE = 0x00000002,  // 8-bit signed int (SBYTE)
            DEVPROP_TYPE_BYTE = 0x00000003,  // 8-bit unsigned int (BYTE)
            DEVPROP_TYPE_INT16 = 0x00000004,  // 16-bit signed int (SHORT)
            DEVPROP_TYPE_UINT16 = 0x00000005,  // 16-bit unsigned int (USHORT)
            DEVPROP_TYPE_INT32 = 0x00000006,  // 32-bit signed int (LONG)
            DEVPROP_TYPE_UINT32 = 0x00000007,  // 32-bit unsigned int (ULONG)
            DEVPROP_TYPE_INT64 = 0x00000008,  // 64-bit signed int (LONG64)
            DEVPROP_TYPE_UINT64 = 0x00000009,  // 64-bit unsigned int (ULONG64)
            DEVPROP_TYPE_FLOAT = 0x0000000A,  // 32-bit floating-point (FLOAT)
            DEVPROP_TYPE_DOUBLE = 0x0000000B,  // 64-bit floating-point (DOUBLE)
            DEVPROP_TYPE_DECIMAL = 0x0000000C,  // 128-bit data (DECIMAL)
            DEVPROP_TYPE_GUID = 0x0000000D,  // 128-bit unique identifier (GUID)
            DEVPROP_TYPE_CURRENCY = 0x0000000E,  // 64 bit signed int currency value (CURRENCY)
            DEVPROP_TYPE_DATE = 0x0000000F,  // date (DATE)
            DEVPROP_TYPE_FILETIME = 0x00000010,  // filetime (FILETIME)
            DEVPROP_TYPE_BOOLEAN = 0x00000011,  // 8-bit boolean (DEVPROP_BOOLEAN)
            DEVPROP_TYPE_STRING = 0x00000012,  // null-terminated string
            DEVPROP_TYPE_STRING_LIST = (DEVPROP_TYPE_STRING | DEVPROP_TYPEMOD_LIST), // multi-sz string list
            DEVPROP_TYPE_SECURITY_DESCRIPTOR = 0x00000013,  // self-relative binary SECURITY_DESCRIPTOR
            DEVPROP_TYPE_SECURITY_DESCRIPTOR_STRING = 0x00000014,  // security descriptor string (SDDL format)
            DEVPROP_TYPE_DEVPROPKEY = 0x00000015,  // device property key (DEVPROPKEY)
            DEVPROP_TYPE_DEVPROPTYPE = 0x00000016,  // device property type (DEVPROPTYPE)
            DEVPROP_TYPE_BINARY = (DEVPROP_TYPE_BYTE | DEVPROP_TYPEMOD_ARRAY),  // custom binary data
            DEVPROP_TYPE_ERROR = 0x00000017,  // 32-bit Win32 system error code
            DEVPROP_TYPE_NTSTATUS = 0x00000018, // 32-bit NTSTATUS code
            DEVPROP_TYPE_STRING_INDIRECT = 0x00000019, // string resource (@[path\]<dllname>,-<strId>)

            MAX_DEVPROP_TYPE = 0x00000019,
            MAX_DEVPROP_TYPEMOD = 0x00002000,

            DEVPROP_MASK_TYPE = 0x00000FFF,
            DEVPROP_MASK_TYPEMOD = 0x0000F000
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct SP_DEVICE_INTERFACE_DATA
        {
            public uint cbSize;
            public Guid InterfaceClassGuid;
            public uint Flags;
            public IntPtr Reserved;
        }

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct SP_DEVINFO_DATA
        {
            public uint cbSize;
            public Guid classGuid;
            public uint devInst;
            public IntPtr reserved;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct DEVPROPKEY
        {
            public Guid fmtid;
            public UInt32 pid;
            public static readonly DEVPROPKEY DEVPKEY_Device_Parent = new DEVPROPKEY { fmtid = new Guid("4340A6C5-93FA-4706-972C-7B648008A5A7"), pid = 8 };
            public static readonly DEVPROPKEY DEVPKEY_Device_Children = new DEVPROPKEY { fmtid = new Guid("4340A6C5-93FA-4706-972C-7B648008A5A7"), pid = 9 };
            public static readonly DEVPROPKEY DEVPKEY_Device_Name = new DEVPROPKEY { fmtid = new Guid("b725f130-47ef-101a-a5f1-02608c9eebac"), pid = 10 };
            public static readonly DEVPROPKEY DEVPKEY_Device_DeviceDesc = new DEVPROPKEY { fmtid = new Guid("a45c254e-df1c-4efd-8020-67d146a850e0"), pid = 2 };
            public static readonly DEVPROPKEY DEVPKEY_Device_FriendlyName = new DEVPROPKEY { fmtid = new Guid("026e516e-b814-414b-83cd-856d6fef4822"), pid = 2 };
            public static readonly DEVPROPKEY DEVPKEY_Device_LocationPaths = new DEVPROPKEY { fmtid = new Guid("a45c254e-df1c-4efd-8020-67d146a850e0"), pid = 37 };
            public static readonly DEVPROPKEY DEVPKEY_Device_LocationInfo = new DEVPROPKEY { fmtid = new Guid("a45c254e-df1c-4efd-8020-67d146a850e0"), pid = 15 };
            public static readonly DEVPROPKEY DEVPKEY_Device_BusReportedDeviceDesc = new DEVPROPKEY { fmtid = new Guid("540b947e-8b40-45bc-a8a2-6a0b894cbda2"), pid = 4 };
            public static readonly DEVPROPKEY DEVPKEY_Device_PDOName = new DEVPROPKEY { fmtid = new Guid("a45c254e-df1c-4efd-8020-67d146a850e0"), pid = 16 };
        }

        [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern bool SetupDiEnumDeviceInterfaces(
        IntPtr deviceInfoSet,
        SP_DEVINFO_DATA deviceInfoData,
        ref Guid interfaceClassGuid,
        int memberIndex,
        SP_DEVICE_INTERFACE_DATA deviceInterfaceData);

        [DllImport("setupapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool SetupDiGetDeviceRegistryProperty(
                IntPtr DeviceInfoSet,
                ref SP_DEVINFO_DATA DeviceInfoData, //ref
                uint Property,
                ref uint PropertyRegDataType,
                IntPtr PropertyBuffer,
                uint PropertyBufferSize,
                ref uint RequiredSize
            );

        [DllImport(@"setupapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool SetupDiDestroyDeviceInfoList(IntPtr hDevInfo);

        [DllImport("setupapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern IntPtr SetupDiGetClassDevs(IntPtr classGuid, string enumerator, IntPtr hwndParent, DiGetClassFlags flags);

        [DllImport("setupapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern bool SetupDiEnumDeviceInfo([In] IntPtr hDevInfo, [In] uint memberIndex, [In, Out] ref SP_DEVINFO_DATA deviceInfoData);

        [DllImport("setupapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool SetupDiGetDeviceProperty(IntPtr DeviceInfoSet, ref SP_DEVINFO_DATA DeviceInfoData, ref DEVPROPKEY propertyKey, ref DEVPROPTYPE propertyType, IntPtr propertyBuffer, uint propertyBufferSize, out uint requiredSize, uint flags);

        [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern bool SetupDiGetDeviceInstanceId(
            IntPtr DeviceInfoSet,
            ref SP_DEVINFO_DATA DeviceInfoData,
            StringBuilder DeviceInstanceId,
            uint DeviceInstanceIdSize,
            out uint RequiredSize
        );

        public SetupApi()
        {
            GetStringProperty(DEVPROPKEY.DEVPKEY_Device_Parent);
        }

        public class Device
        {
            public string FriendlyName;
            public string Parent;
            public string LocationInformation;
            public string LocationPaths;
            public string HardwareId;
            public string Pdo;
            public string DeviceInstancePath;
        }

        public List<Device> devices = new List<Device>();
        public Dictionary<int, string> di = new Dictionary<int, string>();

        public void GetLocationInformation(int DeviceNumber, string ProductName)
        {
            Device AudioLevel1 = null;
            Device AudioLevel2 = null;
            Device AudioLevel3 = null;

            AudioLevel1 = devices.Where(e => e.FriendlyName?.Contains(ProductName) ?? false).FirstOrDefault();
            if (AudioLevel1 != null)
            {
                Console.WriteLine("level 1: {0}", AudioLevel1.FriendlyName);
                Console.WriteLine("level 1: {0}", AudioLevel1.Parent);
                AudioLevel2 = devices.Where(e => e.DeviceInstancePath?.ToUpper().Contains(AudioLevel1.Parent.ToUpper()) ?? false).FirstOrDefault();
                if (AudioLevel2 != null)
                {
                    Console.WriteLine("level 2: {0}", AudioLevel2.FriendlyName);
                    Console.WriteLine("level 2: {0}", AudioLevel2.Parent);
                    AudioLevel3 = devices.Where(e => e.DeviceInstancePath?.ToUpper().Contains(AudioLevel2.Parent.ToUpper()) ?? false).FirstOrDefault();
                    if (AudioLevel3 != null)
                    {
                        Console.WriteLine("level 3: {0}", AudioLevel3.FriendlyName);
                        Console.WriteLine("LocationPaths: {0}", AudioLevel3.LocationPaths);
                        //if (AudioLevel3.LocationInformation.Contains("Hub_"))
                            di.Add(DeviceNumber, AudioLevel3.LocationPaths);
                    }
                }
            }
        }

        public void GetStringProperty(DEVPROPKEY key)
        {
            DEVPROPTYPE dpt = 0;
            uint type = 0;
            uint size = 0;
            bool Success = true;
            uint i = 0;
            int BUFFER_SIZE = 2048;
            IntPtr _hDevInfo = IntPtr.Zero;
            IntPtr buffer = Marshal.AllocHGlobal(BUFFER_SIZE);
            devices.Clear();
            try
            {
                _hDevInfo = SetupDiGetClassDevs(IntPtr.Zero, null, IntPtr.Zero, DiGetClassFlags.DIGCF_PRESENT | DiGetClassFlags.DIGCF_ALLCLASSES);

                if (_hDevInfo == INVALID_HANDLE_VALUE)
                    throw new Win32Exception(Marshal.GetLastWin32Error());

                while (Success)
                {
                    Device device = new Device();
                    SP_DEVINFO_DATA data = new SP_DEVINFO_DATA();
                    data.cbSize = (uint)Marshal.SizeOf(data);
                    Success = SetupDiEnumDeviceInfo(_hDevInfo, i, ref data);

                    if (SetupDiGetDeviceProperty(_hDevInfo, ref data, ref key, ref dpt, buffer, (uint)BUFFER_SIZE, out size, 0))
                    {
                        device.Parent = Marshal.PtrToStringAuto(buffer).Trim();
                    }

                    StringBuilder sb = new StringBuilder(BUFFER_SIZE);
                    if (SetupDiGetDeviceInstanceId(_hDevInfo, ref data, sb, (uint)BUFFER_SIZE, out size))
                    {
                        device.DeviceInstancePath = sb.ToString();
                    }

                    DEVPROPKEY pdokey = DEVPROPKEY.DEVPKEY_Device_PDOName;
                    if (SetupDiGetDeviceProperty(_hDevInfo, ref data, ref pdokey, ref dpt, buffer, (uint)BUFFER_SIZE, out size, 0))
                    {
                        device.Pdo = Marshal.PtrToStringAuto(buffer).Trim();
                    }

                    if (SetupDiGetDeviceRegistryProperty(_hDevInfo, ref data, (int)SPDRP.SPDRP_FRIENDLYNAME, ref type, buffer, (uint)BUFFER_SIZE, ref size))
                    {
                        var s = Marshal.PtrToStringAuto(buffer);
                        device.FriendlyName = s.Trim();
                    }

                    if (SetupDiGetDeviceRegistryProperty(_hDevInfo, ref data, (int)SPDRP.SPDRP_LOCATION_INFORMATION, ref type, buffer, (uint)BUFFER_SIZE, ref size))
                    {
                        var s = Marshal.PtrToStringAuto(buffer);
                        device.LocationInformation = s.Trim();
                    }

                    if (SetupDiGetDeviceRegistryProperty(_hDevInfo, ref data, (int)SPDRP.SPDRP_LOCATION_PATHS, ref type, buffer, (uint)BUFFER_SIZE, ref size))
                    {
                        var s = Marshal.PtrToStringAuto(buffer);
                        device.LocationPaths = s.Trim();
                    }

                    if (SetupDiGetDeviceRegistryProperty(_hDevInfo, ref data, (int)SPDRP.SPDRP_HARDWAREID, ref type, buffer, (uint)BUFFER_SIZE, ref size))
                    {
                        var s = Marshal.PtrToStringAuto(buffer);
                        device.HardwareId = s.Trim();
                    }

                    devices.Add(device);
                    i++;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                SetupDiDestroyDeviceInfoList(_hDevInfo);
                Marshal.FreeHGlobal(buffer);
            }
        }
    }
}
