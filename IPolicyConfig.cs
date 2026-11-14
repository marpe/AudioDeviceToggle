using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace AudioDeviceToggle.Interop;

// Adapted from Vanara's CoreAudio unit tests.
// COM dispatches by vtable slot, so every method must stay declared in this order even if unused.

internal enum DeviceShareMode
{
    Shared,
    Exclusive
}

/// <summary>Undocumented COM-interface IPolicyConfig. Use for set default audio render endpoint.</summary>
[GeneratedComInterface(StringMarshalling = StringMarshalling.Utf16), Guid("f8679f50-850a-41cf-9c72-430f290290c8")]
internal partial interface IPolicyConfig
{
    [PreserveSig]
    int GetMixFormat(string pszDeviceName, out nint ppFormat);

    [PreserveSig]
    int GetDeviceFormat(string pszDeviceName, [MarshalAs(UnmanagedType.Bool)] bool bDefault, out nint ppFormat);

    [PreserveSig]
    int ResetDeviceFormat(string pszDeviceName);

    [PreserveSig]
    int SetDeviceFormat(string pszDeviceName, nint pEndpointFormat, nint mixFormat);

    [PreserveSig]
    int GetProcessingPeriod(string pszDeviceName, [MarshalAs(UnmanagedType.Bool)] bool bDefault, out long pmftDefaultPeriod, out long pmftMinimumPeriod);

    [PreserveSig]
    int SetProcessingPeriod(string pszDeviceName, long pmftPeriod);

    [PreserveSig]
    int GetShareMode(string pszDeviceName, out DeviceShareMode pMode);

    [PreserveSig]
    int SetShareMode(string pszDeviceName, in DeviceShareMode mode);

    [PreserveSig]
    int GetPropertyValue(string pszDeviceName, [MarshalAs(UnmanagedType.Bool)] bool bFxStore, in PROPERTYKEY key, nint pv);

    [PreserveSig]
    int SetPropertyValue(string pszDeviceName, [MarshalAs(UnmanagedType.Bool)] bool bFxStore, in PROPERTYKEY key, nint pv);

    [PreserveSig]
    int SetDefaultEndpoint(string pszDeviceName, ERole role);

    [PreserveSig]
    int SetEndpointVisibility(string pszDeviceName, [MarshalAs(UnmanagedType.Bool)] bool bVisible);
}
