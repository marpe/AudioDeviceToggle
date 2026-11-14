using System.Runtime.InteropServices;
using AudioDeviceToggle.Interop;

namespace AudioDeviceToggle;

internal record AudioDevice(string Id, string Name);

internal static class AudioDevices
{
    // HRESULT_FROM_WIN32(ERROR_NOT_FOUND)
    private const int E_NOTFOUND = unchecked((int)0x80070490);

    private const uint DEVICE_STATE_ACTIVE = 0x1;
    private const uint STGM_READ = 0x0;
    private const ushort VT_LPWSTR = 31;

    private static readonly Guid CLSID_MMDeviceEnumerator = new("BCDE0395-E52F-467C-8E3D-C4579291692E");
    private static readonly Guid CLSID_PolicyConfig = new("870af99c-171d-4f9e-af0d-e63df40c2bc9");

    private static readonly PROPERTYKEY PKEY_Device_FriendlyName = new(new Guid(0xa45c254e, 0xdf1c, 0x4efd, 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0), 14);

    // Same pair Windows sets for "Set as Default Device"; eCommunications is left alone.
    private static readonly ERole[] DefaultRoles = [ERole.eConsole, ERole.eMultimedia];

    public static List<AudioDevice> GetRenderDevices()
    {
        var enumerator = Com.CreateInstance<IMMDeviceEnumerator>(CLSID_MMDeviceEnumerator);
        var collection = enumerator.EnumAudioEndpoints(EDataFlow.eRender, DEVICE_STATE_ACTIVE);

        var count = collection.GetCount();
        var devices = new List<AudioDevice>((int)count);
        for (uint i = 0; i < count; i++)
        {
            devices.Add(ToAudioDevice(collection.Item(i)));
        }

        return devices;
    }

    public static AudioDevice? GetDefault()
    {
        var enumerator = Com.CreateInstance<IMMDeviceEnumerator>(CLSID_MMDeviceEnumerator);
        var hr = enumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out var device);
        if (hr == E_NOTFOUND)
        {
            return null;
        }

        Marshal.ThrowExceptionForHR(hr);
        return ToAudioDevice(device!);
    }

    /// <returns>The HRESULT of the first failing role, or S_OK.</returns>
    public static int SetDefault(string deviceId)
    {
        var policyConfig = Com.CreateInstance<IPolicyConfig>(CLSID_PolicyConfig);
        foreach (var role in DefaultRoles)
        {
            var hr = policyConfig.SetDefaultEndpoint(deviceId, role);
            if (hr < 0)
            {
                return hr;
            }
        }

        return 0;
    }

    private static AudioDevice ToAudioDevice(IMMDevice device)
    {
        return new AudioDevice(device.GetId(), GetFriendlyName(device) ?? string.Empty);
    }

    private static string? GetFriendlyName(IMMDevice device)
    {
        try
        {
            var pv = device.OpenPropertyStore(STGM_READ).GetValue(PKEY_Device_FriendlyName);
            try
            {
                return pv.vt == VT_LPWSTR ? Marshal.PtrToStringUni(pv.pointerValue) : null;
            }
            finally
            {
                Com.PropVariantClear(ref pv);
            }
        }
        catch (COMException)
        {
            return null;
        }
    }
}
