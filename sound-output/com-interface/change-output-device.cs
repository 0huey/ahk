using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;

// based on this answer: https://stackoverflow.com/a/62122685

namespace AudioOutputConsoleApp
{
    public class Program
    {
        public enum HRESULT : uint
        {
            S_OK = 0,
            S_FALSE = 1,
            E_NOINTERFACE = 0x80004002,
            E_NOTIMPL = 0x80004001,
            E_FAIL = 0x80004005,
            E_UNEXPECTED = 0x8000FFFF
        }

        [ComImport]
        [Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public partial interface IPropertyStore
        {
            HRESULT GetCount(out uint propertyCount);
            HRESULT GetAt([In] uint propertyIndex, [MarshalAs(UnmanagedType.Struct)] out PROPERTYKEY key);
            HRESULT GetValue([In][MarshalAs(UnmanagedType.Struct)] ref PROPERTYKEY key, [MarshalAs(UnmanagedType.Struct)] out PROPVARIANT pv);
            HRESULT SetValue([In][MarshalAs(UnmanagedType.Struct)] ref PROPERTYKEY key, [In][MarshalAs(UnmanagedType.Struct)] ref PROPVARIANT pv);
            HRESULT Commit();
        }

        public const int STGM_READ = 0x0;
        public const int STGM_WRITE = 0x1;
        public const int STGM_READWRITE = 0x2;

        public partial struct PROPERTYKEY
        {
            public PROPERTYKEY(Guid InputId, uint InputPid)
            {
                fmtid = InputId;
                pid = InputPid;
            }

            private Guid fmtid;
            private uint pid;
        }

        [StructLayout(LayoutKind.Sequential)]
        public partial struct PROPARRAY
        {
            public uint cElems;
            public IntPtr pElems;
        }

        [StructLayout(LayoutKind.Explicit, Pack = 1)]
        public partial struct PROPVARIANT
        {
            [FieldOffset(0)]
            public ushort varType;
            [FieldOffset(2)]
            public ushort wReserved1;
            [FieldOffset(4)]
            public ushort wReserved2;
            [FieldOffset(6)]
            public ushort wReserved3;
            [FieldOffset(8)]
            public byte bVal;
            [FieldOffset(8)]
            public sbyte cVal;
            [FieldOffset(8)]
            public ushort uiVal;
            [FieldOffset(8)]
            public short iVal;
            [FieldOffset(8)]
            public uint uintVal;
            [FieldOffset(8)]
            public int intVal;
            [FieldOffset(8)]
            public ulong ulVal;
            [FieldOffset(8)]
            public long lVal;
            [FieldOffset(8)]
            public float fltVal;
            [FieldOffset(8)]
            public double dblVal;
            [FieldOffset(8)]
            public short boolVal;
            [FieldOffset(8)]
            public IntPtr pclsidVal;
            [FieldOffset(8)]
            public IntPtr pszVal;
            [FieldOffset(8)]
            public IntPtr pwszVal;
            [FieldOffset(8)]
            public IntPtr punkVal;
            [FieldOffset(8)]
            public PROPARRAY ca;
            [FieldOffset(8)]
            public System.Runtime.InteropServices.ComTypes.FILETIME filetime;
        }

        public enum EDataFlow
        {
            eRender = 0,
            eCapture = eRender + 1,
            eAll = eCapture + 1,
            EDataFlow_enum_count = eAll + 1
        }

        public enum ERole
        {
            eConsole = 0,
            eMultimedia = eConsole + 1,
            eCommunications = eMultimedia + 1,
            ERole_enum_count = eCommunications + 1
        }

        [ComImport]
        [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public partial interface IMMDeviceEnumerator
        {
            HRESULT EnumAudioEndpoints(EDataFlow dataFlow, int dwStateMask, out IMMDeviceCollection ppDevices);
            // for 0x80070490 : Element not found
            [PreserveSig]
            HRESULT GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppEndpoint);
            HRESULT GetDevice(string pwstrId, out IMMDevice ppDevice);
            HRESULT RegisterEndpointNotificationCallback(IMMNotificationClient pClient);
            HRESULT UnregisterEndpointNotificationCallback(IMMNotificationClient pClient);
        }

        public const int DEVICE_STATE_ACTIVE = 0x1;
        public const int DEVICE_STATE_DISABLED = 0x2;
        public const int DEVICE_STATE_NOTPRESENT = 0x4;
        public const int DEVICE_STATE_UNPLUGGED = 0x8;
        public const int DEVICE_STATEMASK_ALL = 0xF;

        [ComImport]
        [Guid("0BD7A1BE-7A1A-44DB-8397-CC5392387B5E")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public partial interface IMMDeviceCollection
        {
            HRESULT GetCount(out uint pcDevices);
            HRESULT Item(uint nDevice, out IMMDevice ppDevice);
        }

        [ComImport]
        [Guid("D666063F-1587-4E43-81F1-B948E807363F")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public partial interface IMMDevice
        {
            HRESULT Activate(ref Guid iid, int dwClsCtx, ref PROPVARIANT pActivationParams, out IntPtr ppInterface);
            HRESULT OpenPropertyStore(int stgmAccess, out IPropertyStore ppProperties);
            HRESULT GetId(out IntPtr ppstrId);
            HRESULT GetState(out int pdwState);
        }

        [ComImport]
        [Guid("7991EEC9-7E89-4D85-8390-6C703CEC60C0")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public partial interface IMMNotificationClient
        {
            HRESULT OnDeviceStateChanged(string pwstrDeviceId, int dwNewState);
            HRESULT OnDeviceAdded(string pwstrDeviceId);
            HRESULT OnDeviceRemoved(string pwstrDeviceId);
            HRESULT OnDefaultDeviceChanged(EDataFlow flow, ERole role, string pwstrDefaultDeviceId);
            HRESULT OnPropertyValueChanged(string pwstrDeviceId, ref PROPERTYKEY key);
        }

        [ComImport]
        [Guid("1BE09788-6894-4089-8586-9A2A6C265AC5")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public partial interface IMMEndpoint
        {
            HRESULT GetDataFlow(out EDataFlow pDataFlow);
        }

        [ComImport]
        [Guid("f8679f50-850a-41cf-9c72-430f290290c8")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public partial interface IPolicyConfig
        {
            HRESULT GetMixFormat([In][MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, out WAVEFORMATEXTENSIBLE ppFormat);
            HRESULT GetDeviceFormat([In][MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [In][MarshalAs(UnmanagedType.Bool)] bool bDefault, out WAVEFORMATEXTENSIBLE ppFormat);
            HRESULT ResetDeviceFormat([In][MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName);
            HRESULT SetDeviceFormat([In][MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [In][MarshalAs(UnmanagedType.LPStruct)] WAVEFORMATEXTENSIBLE pEndpointFormat, [In][MarshalAs(UnmanagedType.LPStruct)] WAVEFORMATEXTENSIBLE pMixFormat);
            HRESULT GetProcessingPeriod([In][MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [In][MarshalAs(UnmanagedType.Bool)] bool bDefault, out long pmftDefaultPeriod, out long pmftMinimumPeriod);
            HRESULT SetProcessingPeriod([In][MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, long pmftPeriod);
            HRESULT GetShareMode([In][MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, out DeviceShareMode pMode);
            HRESULT SetShareMode([In][MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [In] DeviceShareMode mode);
            HRESULT GetPropertyValue([In][MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [In][MarshalAs(UnmanagedType.Bool)] bool bFxStore, ref PROPERTYKEY pKey, out PROPVARIANT pv);
            HRESULT SetPropertyValue([In][MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [In][MarshalAs(UnmanagedType.Bool)] bool bFxStore, [In] ref PROPERTYKEY pKey, ref PROPVARIANT pv);
            HRESULT SetDefaultEndpoint([In][MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [In][MarshalAs(UnmanagedType.U4)] ERole role);
            HRESULT SetEndpointVisibility([In][MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [In][MarshalAs(UnmanagedType.Bool)] bool bVisible);
        }

        [StructLayout(LayoutKind.Explicit, Pack = 1)]
        public partial class WAVEFORMATEXTENSIBLE : WAVEFORMATEX
        {
            [FieldOffset(0)]
            public short wValidBitsPerSample;
            [FieldOffset(0)]
            public short wSamplesPerBlock;
            [FieldOffset(0)]
            public short wReserved;
            [FieldOffset(2)]
            public WaveMask dwChannelMask;
            [FieldOffset(6)]
            public Guid SubFormat;
        }

        [Flags]
        public enum WaveMask
        {
            None = 0x0,
            FrontLeft = 0x1,
            FrontRight = 0x2,
            FrontCenter = 0x4,
            LowFrequency = 0x8,
            BackLeft = 0x10,
            BackRight = 0x20,
            FrontLeftOfCenter = 0x40,
            FrontRightOfCenter = 0x80,
            BackCenter = 0x100,
            SideLeft = 0x200,
            SideRight = 0x400,
            TopCenter = 0x800,
            TopFrontLeft = 0x1000,
            TopFrontCenter = 0x2000,
            TopFrontRight = 0x4000,
            TopBackLeft = 0x8000,
            TopBackCenter = 0x10000,
            TopBackRight = 0x20000
        }

        [StructLayout(LayoutKind.Sequential, Pack = 2)]
        public partial class WAVEFORMATEX
        {
            public short wFormatTag;
            public short nChannels;
            public int nSamplesPerSec;
            public int nAvgBytesPerSec;
            public short nBlockAlign;
            public short wBitsPerSample;
            public short cbSize;
        }

        public enum DeviceShareMode
        {
            Shared,
            Exclusive
        }

        private static PROPERTYKEY PKEY_Device_FriendlyName = new PROPERTYKEY(new Guid("a45c254e-df1c-4efd-8020-67d146a850e0"), 14);
        private static PROPERTYKEY PKEY_Device_DeviceDesc = new PROPERTYKEY(new Guid("a45c254e-df1c-4efd-8020-67d146a850e0"), 2);
        private static PROPERTYKEY PKEY_DeviceClass_IconPath = new PROPERTYKEY(new Guid("259abffc-50a7-47ce-af08-68c9a7d73366"), 12);

        private enum RoleArgSwitch
        {
            console = 0,
            multimedia = console + 1,
            communications = multimedia + 1,
            all = communications + 1
        }

        public static void Main(string[] args)
        {
            bool PrintActiveDevices = false;
            RoleArgSwitch RoleSwitch = RoleArgSwitch.all;
            string DeviceName = "";

            if (args.Length == 0)
            {
                PrintActiveDevices = true;
            }
            else if (args.Length == 2)
            {
                if (args[0] == "--console")
                {
                    RoleSwitch = RoleArgSwitch.console;
                }
                else if (args[0] == "--multimedia")
                {
                    RoleSwitch = RoleArgSwitch.multimedia;
                }
                else if (args[0] == "--communications")
                {
                    RoleSwitch = RoleArgSwitch.communications;
                }
                else if (args[0] == "--all")
                {
                    RoleSwitch = RoleArgSwitch.all;
                }
                else
                {
                    PrintHelp();
                    return;
                }

                DeviceName = args[1].ToLower();
            }
            else
            {
                PrintHelp();
                return;
            }

            var devices = GetDevices();

            foreach (Dictionary<string, string> device in devices)
            {
                if (PrintActiveDevices)
                {
                    Console.WriteLine(device["name"]);
                }
                else if (device["name"] == DeviceName)
                {
                    SetDefaultOutput(device["id"], RoleSwitch);
                    return;
                }
            }
            if (!PrintActiveDevices) {
                Console.WriteLine("Device name not found");
            }
        }

        private static List<Dictionary<string, string>> GetDevices()
        {
            var devices = new List<Dictionary<string, string>>();
            var hr = HRESULT.E_FAIL;

            var CLSID_MMDeviceEnumerator = new Guid("{BCDE0395-E52F-467C-8E3D-C4579291692E}");
            var MMDeviceEnumeratorType = Type.GetTypeFromCLSID(CLSID_MMDeviceEnumerator, true);
            var MMDeviceEnumerator = Activator.CreateInstance(MMDeviceEnumeratorType);
            IMMDeviceEnumerator pMMDeviceEnumerator = (IMMDeviceEnumerator)MMDeviceEnumerator;
            if (pMMDeviceEnumerator == null)
            {
                Console.WriteLine("error creating DeviceEnumerator");
                return devices;
            }

            IMMDeviceCollection pDeviceCollection = null;
            hr = pMMDeviceEnumerator.EnumAudioEndpoints(EDataFlow.eRender, DEVICE_STATE_ACTIVE, out pDeviceCollection);
            if (hr != HRESULT.S_OK)
            {
                Console.WriteLine("Error enumerating devices");
                return devices;
            }

            uint nDevices = 0;
            hr = pDeviceCollection.GetCount(out nDevices);
            if (nDevices == 0) //Return when no devices are enumerated.
            {
                return devices;
            }

            for (uint i = 0, loopTo = (uint)(nDevices - (long)1); i <= loopTo; i++)
            {
                IMMDevice pDevice = null;
                hr = pDeviceCollection.Item(i, out pDevice);
                if (hr != HRESULT.S_OK)
                {
                    Console.WriteLine("Error getting device index: " + i);
                    continue;
                }

                IPropertyStore pPropertyStore = null;
                hr = pDevice.OpenPropertyStore(STGM_READ, out pPropertyStore);
                if (hr != HRESULT.S_OK)
                {
                    Console.WriteLine("Error getting property store for device index: " + i);
                    continue;
                }

                string sFriendlyName = null;
                var pv = new PROPVARIANT();
                hr = pPropertyStore.GetValue(ref PKEY_Device_FriendlyName, out pv);
                if (hr != HRESULT.S_OK)
                {
                    Console.WriteLine("Error getting name for device index: " + i);
                    continue;
                }

                sFriendlyName = Marshal.PtrToStringUni(pv.pwszVal);

                // remove unnecessary trailing info: '(High Definition Audio Device)'
                sFriendlyName = sFriendlyName.Split('(')[0];
                sFriendlyName = sFriendlyName.ToLower();
                sFriendlyName = sFriendlyName.Trim();

                var hGlobal = Marshal.AllocHGlobal(260);
                hr = pDevice.GetId(out hGlobal);
                string sId = Marshal.PtrToStringUni(hGlobal);
                Marshal.FreeHGlobal(hGlobal);

                if (hr != HRESULT.S_OK)
                {
                    Console.WriteLine("Error getting id for device index: " + i);
                    continue;
                }

                var data = new Dictionary<string, string>();
                data.Add("name", sFriendlyName);
                data.Add("id", sId);
                devices.Add(data);
            }
            return devices;
        }

        private static void SetDefaultOutput(string Id, RoleArgSwitch output_type)
        {
            var hr = HRESULT.E_FAIL;
            var CLSID_PolicyConfig = new Guid("{870af99c-171d-4f9e-af0d-e63df40c2bc9}");
            var PolicyConfigType = Type.GetTypeFromCLSID(CLSID_PolicyConfig, true);
            var PolicyConfig = Activator.CreateInstance(PolicyConfigType);
            IPolicyConfig pPolicyConfig = (IPolicyConfig)PolicyConfig;

            if (pPolicyConfig is object)
            {
                if (output_type == RoleArgSwitch.console || output_type == RoleArgSwitch.all)
                {
                    hr = pPolicyConfig.SetDefaultEndpoint(Id, ERole.eConsole);
                }

                if (output_type == RoleArgSwitch.multimedia || output_type == RoleArgSwitch.all)
                {
                    hr = pPolicyConfig.SetDefaultEndpoint(Id, ERole.eMultimedia);
                }

                if (output_type == RoleArgSwitch.communications || output_type == RoleArgSwitch.all)
                {
                    hr = pPolicyConfig.SetDefaultEndpoint(Id, ERole.eCommunications);
                }

                if (hr != HRESULT.S_OK)
                {
                    Console.WriteLine("Error setting output device");
                }

                Marshal.ReleaseComObject(PolicyConfig);
            }
        }

        private static void PrintHelp()
        {
            string prog_name = typeof(Program).Assembly.GetModules()[0].Name;

            Console.WriteLine("Usage: {0} [OUTPUT_TYPE DEVICE_NAME]", prog_name);
            Console.WriteLine("No arguments will list the available output device names");
            Console.WriteLine("the available OUTPUT_TYPE values are:");
            Console.WriteLine("    --console");
            Console.WriteLine("    --multimedia");
            Console.WriteLine("    --communications");
            Console.WriteLine("    --all");
        }
    }
 }
