using System.Runtime.InteropServices;
using AudioDeviceToggle;

public enum DeviceShareMode
{
    Shared,
    Exclusive
}

/// <summary>Undocumented COM-interface IPolicyConfig. Use for set default audio render endpoint.</summary>
[ComImport, Guid("f8679f50-850a-41cf-9c72-430f290290c8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), CoClass(typeof(PolicyConfig))]
internal interface IPolicyConfig
{
    [PreserveSig]
    int GetMixFormat([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [Out] out IntPtr ppFormat);

    [PreserveSig]
    int GetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [MarshalAs(UnmanagedType.Bool)] bool bDefault, [Out] out IntPtr ppFormat);

    [PreserveSig]
    int ResetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName);

    [PreserveSig]
    int SetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [In] IntPtr pEndpointFormat, [In] IntPtr mixFormat);

    [PreserveSig]
    int GetProcessingPeriod([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [MarshalAs(UnmanagedType.Bool)] bool bDefault, out long pmftDefaultPeriod, out long pmftMinimumPeriod);

    [PreserveSig]
    int SetProcessingPeriod([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, long pmftPeriod);

    [PreserveSig]
    int GetShareMode([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [Out] out DeviceShareMode pMode);

    [PreserveSig]
    int SetShareMode([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, in DeviceShareMode mode);

    [PreserveSig]
    int GetPropertyValue([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [MarshalAs(UnmanagedType.Bool)] bool bFxStore, in PROPERTYKEY key, [Out] PropVariant pv);

    [PreserveSig]
    int SetPropertyValue([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [MarshalAs(UnmanagedType.Bool)] bool bFxStore, in PROPERTYKEY key, [Out] PropVariant pv);

    [PreserveSig]
    int SetDefaultEndpoint([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [MarshalAs(UnmanagedType.U4)] ERole role);

    [PreserveSig]
    int SetEndpointVisibility([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, [MarshalAs(UnmanagedType.Bool)] bool bVisible);
}

[ComImport, Guid("870af99c-171d-4f9e-af0d-e63df40c2bc9"), ClassInterface(ClassInterfaceType.None)]
public class PolicyConfig
{
}

internal enum ERole
{
    eConsole = 0,
    eMultimedia = 1,
    eCommunications = 2
}

