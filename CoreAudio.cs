using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace AudioDeviceToggle.Interop;

// Minimal Core Audio (mmdeviceapi.h) and property store bindings. Source-generated COM ([GeneratedComInterface])
// rather than [ComImport], because built-in COM interop is unavailable under Native AOT.
// COM dispatches by vtable slot, so methods stay in header order; trailing methods that are never called are omitted.

internal enum EDataFlow
{
    eRender,
    eCapture,
    eAll
}

internal enum ERole
{
    eConsole,
    eMultimedia,
    eCommunications
}

[StructLayout(LayoutKind.Sequential)]
internal struct PROPERTYKEY(Guid fmtid, uint pid)
{
    public Guid fmtid = fmtid;
    public uint pid = pid;
}

// Only the header and first pointer-sized union member are read; the trailing field pads the struct to its native size.
[StructLayout(LayoutKind.Sequential)]
internal struct PROPVARIANT
{
    public ushort vt;
    private ushort reserved1;
    private ushort reserved2;
    private ushort reserved3;
    public nint pointerValue;
    private nint padding;
}

[GeneratedComInterface(StringMarshalling = StringMarshalling.Utf16), Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
internal partial interface IMMDeviceEnumerator
{
    IMMDeviceCollection EnumAudioEndpoints(EDataFlow dataFlow, uint stateMask);

    [PreserveSig]
    int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice? device);
}

[GeneratedComInterface, Guid("0BD7A1BE-7A1A-44DB-8397-CC5392387B5E")]
internal partial interface IMMDeviceCollection
{
    uint GetCount();

    IMMDevice Item(uint index);
}

[GeneratedComInterface(StringMarshalling = StringMarshalling.Utf16), Guid("D666063F-1587-4E43-81F1-B948E807363F")]
internal partial interface IMMDevice
{
    // Never called; declared so the following methods land in the right vtable slots.
    void Activate(nint iid, uint clsCtx, nint activationParams, out nint instance);

    IPropertyStore OpenPropertyStore(uint stgmAccess);

    string GetId();
}

[GeneratedComInterface, Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99")]
internal partial interface IPropertyStore
{
    uint GetCount();

    PROPERTYKEY GetAt(uint index);

    PROPVARIANT GetValue(in PROPERTYKEY key);
}

internal static partial class Com
{
    private const uint CLSCTX_ALL = 0x17;

    private static readonly StrategyBasedComWrappers ComWrappers = new();

    public static T CreateInstance<T>(Guid clsid) where T : class
    {
        var iid = StrategyBasedComWrappers.DefaultIUnknownInterfaceDetailsStrategy.GetIUnknownDerivedDetails(typeof(T).TypeHandle)!.Iid;
        Marshal.ThrowExceptionForHR(CoCreateInstance(clsid, 0, CLSCTX_ALL, iid, out var pointer));
        try
        {
            return (T)ComWrappers.GetOrCreateObjectForComInstance(pointer, CreateObjectFlags.UniqueInstance);
        }
        finally
        {
            Marshal.Release(pointer);
        }
    }

    [LibraryImport("ole32.dll")]
    private static partial int CoCreateInstance(in Guid clsid, nint outer, uint clsContext, in Guid iid, out nint instance);

    [LibraryImport("ole32.dll")]
    public static partial int PropVariantClear(ref PROPVARIANT pv);
}
