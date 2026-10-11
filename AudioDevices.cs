using System.Runtime.InteropServices;
using AudioDeviceToggle.Interop;
using Vanara.InteropServices;
using Vanara.PInvoke;
using static Vanara.PInvoke.CoreAudio;

namespace AudioDeviceToggle;

internal record AudioDevice(string Id, string Name);

/// <summary>
/// Core Audio COM access for render (output) devices. COM objects are released before returning,
/// so callers only ever see plain <see cref="AudioDevice"/> records.
/// </summary>
// https://github.com/dahall/Vanara/blob/67dc6e5e42cbd08b5a3dd9ff5c3ebd72283876a5/UnitTests/PInvoke/CoreAudio/DeviceTests.cs
internal static class AudioDevices
{
    // HRESULT_FROM_WIN32(ERROR_NOT_FOUND): returned when no device exists for the requested flow/role.
    private const int E_NOTFOUND = unchecked((int)0x80070490);

    private static readonly Ole32.PROPERTYKEY PKEY_Device_FriendlyName = new(new Guid(0xa45c254e, 0xdf1c, 0x4efd, 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0), 14);

    // Same pair Windows sets for "Set as Default Device"; eCommunications is left alone.
    private static readonly ERole[] DefaultRoles = [ERole.eConsole, ERole.eMultimedia];

    public static List<AudioDevice> GetRenderDevices()
    {
        using var enumerator = ComReleaserFactory.Create(new IMMDeviceEnumerator());
        using var collection = ComReleaserFactory.Create(enumerator.Item.EnumAudioEndpoints(EDataFlow.eRender, DEVICE_STATE.DEVICE_STATE_ACTIVE)!);

        var count = collection.Item.GetCount();
        var devices = new List<AudioDevice>((int)count);
        for (uint i = 0; i < count; i++)
        {
            collection.Item.Item(i, out var device).ThrowIfFailed();
            using var deviceReleaser = ComReleaserFactory.Create(device!);
            devices.Add(ToAudioDevice(deviceReleaser.Item));
        }

        return devices;
    }

    /// <returns>The default render device, or <c>null</c> if no output device is active.</returns>
    public static AudioDevice? GetDefault()
    {
        using var enumerator = ComReleaserFactory.Create(new IMMDeviceEnumerator());
        IMMDevice device;
        try
        {
            device = enumerator.Item.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia)!;
        }
        catch (COMException ex) when (ex.HResult == E_NOTFOUND)
        {
            return null;
        }

        using var deviceReleaser = ComReleaserFactory.Create(device);
        return ToAudioDevice(deviceReleaser.Item);
    }

    public static HRESULT SetDefault(string deviceId)
    {
        using var policyConfig = ComReleaserFactory.Create(new IPolicyConfig());
        foreach (var role in DefaultRoles)
        {
            var hr = policyConfig.Item.SetDefaultEndpoint(deviceId, role);
            if (hr.Failed)
            {
                return hr;
            }
        }

        return HRESULT.S_OK;
    }

    private static AudioDevice ToAudioDevice(IMMDevice device)
    {
        return new AudioDevice(device.GetId()!, GetFriendlyName(device) ?? string.Empty);
    }

    private static string? GetFriendlyName(IMMDevice device)
    {
        try
        {
            using var props = ComReleaserFactory.Create(device.OpenPropertyStore(STGM.STGM_READ)!);
            using var pv = new Ole32.PROPVARIANT();
            props.Item.GetValue(PKEY_Device_FriendlyName, pv);
            return pv.pwszVal;
        }
        catch (COMException)
        {
            return null;
        }
    }
}
