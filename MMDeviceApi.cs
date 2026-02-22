/*
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace AudioDeviceToggle
{
    internal static class MMConstants
    {
        internal const uint STGM_READ = 0;
        internal const uint STGM_WRITE = 1;
        internal const uint STGM_READWRITE = 2;
        internal static readonly PROPERTYKEY PKEY_Device_FriendlyName = new("{A45C254E-DF1C-4EFD-8020-67D146A850E0}", 0xE);
        internal static readonly PROPERTYKEY PKEY_AudioEndpoint_Disable_SysFx = new("{1da5d803-d492-4edd-8c23-e0c0ffee7f0e}", 5);
    }

    [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    internal class MMDeviceEnumerator
    {
    }

    [ComImport, Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDeviceEnumerator
    {
        [PreserveSig]
        int EnumAudioEndpoints(EDataFlow dataFlow, DeviceState dwStateMask, out IMMDeviceCollection ppDevices);
        [PreserveSig]
        int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppEndpoint);
        [PreserveSig]
        int GetDevice(string pwstrId, out IMMDevice ppDevice);
        [PreserveSig]
        int RegisterEndpointNotificationCallback(IntPtr pClient);
        [PreserveSig]
        int UnregisterEndpointNotificationCallback(IntPtr pClient);
    }

    [ComImport, Guid("0BD7A1BE-7A1A-44DB-8397-CC5392387B5E"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDeviceCollection
    {
        [PreserveSig]
        int GetCount(out uint pcDevices);

        [PreserveSig]
        int Item(uint nDevice, out IMMDevice ppDevice);
    }

    [ComImport, Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDevice
    {
        [PreserveSig]
        int Activate();

        [PreserveSig]
        int OpenPropertyStore(uint stgmAccess, out IPropertyStore ppProperties);

        [PreserveSig]
        int GetId([MarshalAs(UnmanagedType.LPWStr)] out string id);

        [PreserveSig]
        int GetState(out DeviceState state);
    }

    [ComImport, Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IPropertyStore
    {
        [PreserveSig]
        int GetCount(out uint cProps);

        [PreserveSig]
        int GetAt(uint iProp, out PROPERTYKEY pkey);

        [PreserveSig]
        int GetValue(ref PROPERTYKEY key, out PropVariant pv);

        [PreserveSig]
        int SetValue(ref PROPERTYKEY key, ref PropVariant propvar);

        [PreserveSig]
        int Commit();
    }

    [ComImport, Guid("67c5fc9c-29e1-4154-8307-84ed8edb5a21"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface ILocalEndpointDevice
    {
        [PreserveSig]
        int IsEnabled([MarshalAs(UnmanagedType.Bool)] out bool bEnabled);

        [PreserveSig]
        int GetInternalState(out uint pdwState);

        [PreserveSig]
        int OpenFXPropertyStore(uint stgmAccess, out IPropertyStore ppFxProperties);
    }

    internal enum EDataFlow : uint
    {
        eRender = 0,
        eCapture,
        eAll,
        EDataFlow_enum_count,
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct PROPERTYKEY
    {
        internal Guid fmtid;
        internal uint pid;

        internal PROPERTYKEY(string fmtidStr, uint pid)
        {
            fmtid = Guid.Parse(fmtidStr);
            this.pid = pid;
        }
    }

    [StructLayout(LayoutKind.Sequential, Pack = 0)]
    public struct PropArray
    {
        internal uint cElems;
        internal IntPtr pElems;
    }

    [StructLayout(LayoutKind.Explicit, Pack = 1)]
    internal struct PropVariant
    {
        [FieldOffset(0)]
        internal VarEnum varType;

        [FieldOffset(2)]
        internal ushort wReserved1;

        [FieldOffset(4)]
        internal ushort wReserved2;

        [FieldOffset(6)]
        internal ushort wReserved3;

        [FieldOffset(8)]
        internal byte bVal;

        [FieldOffset(8)]
        internal sbyte cVal;

        [FieldOffset(8)]
        internal ushort uiVal;

        [FieldOffset(8)]
        internal short iVal;

        [FieldOffset(8)]
        internal uint uintVal;

        [FieldOffset(8)]
        internal int intVal;

        [FieldOffset(8)]
        internal ulong ulVal;

        [FieldOffset(8)]
        internal long lVal;

        [FieldOffset(8)]
        internal float fltVal;

        [FieldOffset(8)]
        internal double dblVal;

        [FieldOffset(8)]
        internal short boolVal;

        [FieldOffset(8)]
        internal IntPtr pclsidVal;

        [FieldOffset(8)]
        internal IntPtr pszVal;

        [FieldOffset(8)]
        internal IntPtr pwszVal;

        [FieldOffset(8)]
        internal IntPtr punkVal;

        [FieldOffset(8)]
        internal PropArray ca;

        [FieldOffset(8)]
        internal FILETIME filetime;
    }
}
*/
