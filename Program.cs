using System.Runtime.InteropServices;
using Vanara.InteropServices;
using Vanara.PInvoke;

// https://github.com/dahall/Vanara/blob/67dc6e5e42cbd08b5a3dd9ff5c3ebd72283876a5/UnitTests/PInvoke/CoreAudio/DeviceTests.cs

// -c "{0.0.0.00000000}.{d0d43511-8c68-4b84-a640-f994f4903609}" "{0.0.0.00000000}.{0dc6ae6b-03e1-43cb-99b7-fec1bda2b5b2}"

Ole32.PROPERTYKEY PKEY_Device_FriendlyName = new(new Guid(0xa45c254e, 0xdf1c, 0x4efd, 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0), 14);

var commands = new List<CommandLineArg>
{
    new("list", ["list", "l"], (_) => ListDevices()),
    new("set device by id", ["id", "i"], HandleSetDefaultDeviceById),
    new("set device by name", ["name", "n"], HandleSetDefaultDeviceByName),
    new("cycle devices by id", ["ci"], CycleDevicesById),
    new("cycle devices by name", ["c"], CycleDevicesByName),
};

commands.Add(new CommandLineArg("help", ["?", "h", "help"], (_) => PrintUsage()));

foreach (var arg in args)
{
    foreach (var cmd in commands)
    {
        if (MatchesArgument(arg, cmd))
        {
            cmd.Callback(args);
            return;
        }
    }
}

PrintUsage();
Console.WriteLine();
ListDevices();

return;

IEnumerable<CoreAudio.IMMDevice> CreateIMMDeviceCollection(CoreAudio.IMMDeviceEnumerator deviceEnumerator, CoreAudio.EDataFlow direction = CoreAudio.EDataFlow.eAll, CoreAudio.DEVICE_STATE stateMasks = CoreAudio.DEVICE_STATE.DEVICE_STATEMASK_ALL)
{
    using var deviceCollection = ComReleaserFactory.Create(deviceEnumerator.EnumAudioEndpoints(direction, stateMasks)!);
    var deviceList = new List<CoreAudio.IMMDevice>();
    var cnt = deviceCollection.Item.GetCount();
    for (uint i = 0; i < cnt; i++)
    {
        deviceCollection.Item.Item(i, out var dev).ThrowIfFailed();
        deviceList.Add(dev!);
    }

    return deviceList;
}

CoreAudio.IMMDevice? GetDefaultAudioEndpoint(CoreAudio.EDataFlow dataFlow, CoreAudio.ERole role)
{
    using var enumerator = ComReleaserFactory.Create(new CoreAudio.IMMDeviceEnumerator());
    var device = enumerator.Item.GetDefaultAudioEndpoint(dataFlow, role);
    return device;
}

List<CoreAudio.IMMDevice> GetDevices(CoreAudio.EDataFlow flow, CoreAudio.DEVICE_STATE stateMask = CoreAudio.DEVICE_STATE.DEVICE_STATE_ACTIVE)
{
    using var enumerator = ComReleaserFactory.Create(new CoreAudio.IMMDeviceEnumerator());
    return CreateIMMDeviceCollection(enumerator.Item, flow, stateMask).ToList();
}

void ListDevices()
{
    var flow = CoreAudio.EDataFlow.eRender;
    var role = CoreAudio.ERole.eMultimedia;

    var devices = GetDevices(flow);

    var current = GetDefaultAudioEndpoint(flow, role);
    var currentId = current?.GetId()! ?? string.Empty;

    var defaultForegroundColor = Console.ForegroundColor;
    Console.WriteLine("Device list:");

    for (var i = 0; i < devices.Count; i++)
    {
        var device = devices[i];

        var name = GetDeviceName(device);
        var deviceId = device.GetId();
        var state = device.GetState();

        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.Write($"{i + 1}.");

        var isCurrent = currentId != null && currentId.Equals(deviceId, StringComparison.OrdinalIgnoreCase);

        Console.ForegroundColor = defaultForegroundColor;
        Console.Write(" \"{0}\"", name);

        Console.ForegroundColor = isCurrent ? ConsoleColor.Red : defaultForegroundColor;
        Console.WriteLine(isCurrent ? " (Current Default)" : "");
    }
}

void PrintUsage()
{
    Console.WriteLine("Usage:");

    foreach (var command in commands)
    {
        var defaultForegroundColor = Console.ForegroundColor;
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine($"{command.Name} ({string.Join(", ", command.Aliases)})");
        Console.ForegroundColor = defaultForegroundColor;
    }
}


void CycleDevicesByName(string[] args)
{
    if (args.Length < 3)
    {
        Console.WriteLine("At least two device names must be specified to cycle through.");
        return;
    }

    var role = CoreAudio.ERole.eMultimedia;
    var flow = CoreAudio.EDataFlow.eRender;
    var defaultDevice = GetDefaultAudioEndpoint(flow, role);
    var defaultDeviceName = string.Empty;
    var indexOfCurrentDefaultDevice = -1;

    if (defaultDevice != null)
    {
        defaultDeviceName = GetDeviceName(defaultDevice);

        if (!string.IsNullOrEmpty(defaultDeviceName))
        {
            for (var i = 1; i < args.Length; i++)
            {
                if (defaultDeviceName.Equals(args[i], StringComparison.OrdinalIgnoreCase))
                {
                    indexOfCurrentDefaultDevice = i - 1;
                    break;
                }
            }
        }
    }

    var nextDeviceName = args[1 + ((indexOfCurrentDefaultDevice + 1) % (args.Length - 1))];

    var defaultForegroundColor = Console.ForegroundColor;
    Console.WriteLine($"Switching ({role}) device");
    Console.Write($"from: ");
    Console.ForegroundColor = ConsoleColor.Magenta;
    Console.WriteLine(defaultDeviceName);
    Console.ForegroundColor = defaultForegroundColor;
    Console.Write("to: ");
    Console.ForegroundColor = ConsoleColor.Magenta;
    Console.WriteLine(nextDeviceName);
    Console.ForegroundColor = defaultForegroundColor;

    SetDefaultDeviceByName(nextDeviceName);
}


void CycleDevicesById(string[] args)
{
    if (args.Length < 3)
    {
        Console.WriteLine("At least two device ids must be specified to cycle through.");
        return;
    }

    var role = CoreAudio.ERole.eMultimedia;
    var flow = CoreAudio.EDataFlow.eRender;
    var defaultDevice = GetDefaultAudioEndpoint(flow, role);
    var defaultDeviceId = string.Empty;
    var indexOfCurrentDefaultDevice = -1;

    if (defaultDevice != null)
    {
        defaultDeviceId = defaultDevice.GetId()!;

        if (!string.IsNullOrEmpty(defaultDeviceId))
        {
            for (var i = 1; i < args.Length; i++)
            {
                if (defaultDeviceId.Equals(args[i], StringComparison.OrdinalIgnoreCase))
                {
                    indexOfCurrentDefaultDevice = i - 1;
                    break;
                }
            }
        }
    }

    var nextDeviceId = args[1 + ((indexOfCurrentDefaultDevice + 1) % (args.Length - 1))];

    Console.WriteLine($"Switching ({role}) device from id: {defaultDeviceId} to id: {nextDeviceId}");

    SetDefaultDeviceById(nextDeviceId, role);
}

void HandleSetDefaultDeviceById(string[] args)
{
    if (args.Length < 2)
    {
        Console.WriteLine("Device id not specified.");
        return;
    }

    var newDeviceId = args[1];
    var role = CoreAudio.ERole.eMultimedia;
    SetDefaultDeviceById(newDeviceId, role);
}

void HandleSetDefaultDeviceByName(string[] args)
{
    if (args.Length < 2)
    {
        Console.WriteLine("Device name not specified.");
        return;
    }

    var deviceName = args[1];
    SetDefaultDeviceByName(deviceName);
}

void SetDefaultDeviceByName(string deviceName)
{
    if (args.Length < 2)
    {
        Console.WriteLine("Device name not specified.");
        return;
    }

    var flow = CoreAudio.EDataFlow.eRender;
    var devices = GetDevices(flow);

    var device = devices.FirstOrDefault(d =>
    {
        var name = GetDeviceName(d);

        if (string.IsNullOrEmpty(name))
        {
            Console.WriteLine($"Device has an empty name, skipping.");
            return false;
        }

        return name.Equals(deviceName, StringComparison.OrdinalIgnoreCase);
    });

    if (device == null)
    {
        PrintError($"Could not find device with name: {deviceName}");
        ListDevices();
        return;
    }

    var deviceId = device.GetId()!;
    var role = CoreAudio.ERole.eMultimedia;

    SetDefaultDeviceById(deviceId, role);
}

void SetDefaultDeviceById(string newDeviceId, CoreAudio.ERole role)
{
    var flow = Vanara.PInvoke.CoreAudio.EDataFlow.eRender;
    using var policyConfig = ComReleaserFactory.Create(new Vanara.PInvoke.Tests.CoreAudio.IPolicyConfig());
    var devices = GetDevices(flow);

    var newDevice = devices.FirstOrDefault(d =>
    {
        var id = d.GetId()!;
        return id.ToString()!.Equals(newDeviceId, StringComparison.OrdinalIgnoreCase);
    });

    if (newDevice == null)
    {
        Console.WriteLine($"Could not find device with id: {newDeviceId}");
        return;
    }

    var hr = policyConfig.Item.SetDefaultEndpoint(newDeviceId, role);
    if (PrintIfFailed(hr.Code, $"Failed to set default endpoint for role {role}"))
    {
        return;
    }

    var deviceName = GetDeviceName(newDevice);
    var defaultForegroundColor = Console.ForegroundColor;
    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.Write($"Default endpoint changed to ");
    Console.ForegroundColor = ConsoleColor.White;
    Console.WriteLine($" \"{deviceName}\"");
    Console.ForegroundColor = defaultForegroundColor;
    Console.WriteLine();

    ListDevices();
}

bool MatchesArgument(string arg, CommandLineArg command)
{
    return Array.Exists(command.Aliases, a => a.Equals(arg, StringComparison.OrdinalIgnoreCase) ||
                                              $"/{a}".Equals(arg, StringComparison.OrdinalIgnoreCase) ||
                                              $"-{a}".Equals(arg, StringComparison.OrdinalIgnoreCase) ||
                                              $"--{a}".Equals(arg, StringComparison.OrdinalIgnoreCase));
}


void PrintError(string message, int hResult = 0)
{
    var defaultForegroundColor = Console.ForegroundColor;
    Console.ForegroundColor = ConsoleColor.Red;
    Console.Write($"Error: ");
    Console.ForegroundColor = ConsoleColor.White;
    Console.WriteLine($"{message}{(hResult != 0 ? $" (hr = 0x{hResult:X8})" : string.Empty)}\n");
    Console.ForegroundColor = defaultForegroundColor;
}

/*
 * Any COM method that returns an HRESULT will return a negative value if it failed. This helper method prints the error message along with the HRESULT in hex format if the call failed, and returns true if it failed.
 */
bool PrintIfFailed(int hResult, string message)
{
    if (hResult >= 0)
    {
        return false;
    }

    PrintError(message, hResult);
    return true;
}

string? GetDeviceNameById(string devId)
{
    using var pEnum = ComReleaserFactory.Create(new CoreAudio.IMMDeviceEnumerator());
    using var pDev = ComReleaserFactory.Create(pEnum.Item.GetDevice(devId)!);
    using var pProps = ComReleaserFactory.Create(pDev.Item.OpenPropertyStore(STGM.STGM_READ)!);
    using var pv = new Ole32.PROPVARIANT();
    try
    {
        pProps.Item.GetValue(PKEY_Device_FriendlyName, pv);
        return pv.pwszVal;
    }
    catch
    {
    }

    return null;
}

string? GetDeviceName(CoreAudio.IMMDevice mmDevice)
{
    return GetDeviceNameById(mmDevice.GetId()!);
}

internal record CommandLineArg(string Name, string[] Aliases, Action<string[]> Callback)
{
    public string Name = Name;
    public string[] Aliases = Aliases;
    public Action<string[]> Callback = Callback;
}
