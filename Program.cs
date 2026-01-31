using System.Runtime.InteropServices;
using AudioDeviceToggle;

// https://github.com/dahall/Vanara/blob/67dc6e5e42cbd08b5a3dd9ff5c3ebd72283876a5/UnitTests/PInvoke/CoreAudio/DeviceTests.cs

// -c "{0.0.0.00000000}.{d0d43511-8c68-4b84-a640-f994f4903609}" "{0.0.0.00000000}.{0dc6ae6b-03e1-43cb-99b7-fec1bda2b5b2}"

var commands = new List<CommandLineArg>
{
    new("list", ["list", "l"], ListDevices),
    new("set default device by id", ["id", "i"], SetDefaultDeviceById),
    new("set default device by name", ["name", "n"], SetDefaultDeviceByName),
    new("cycle devices", ["c"], CycleDevices)
};

commands.Add(new CommandLineArg("help", ["?", "h", "help"], PrintUsage));

ListDevices(args);

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

return;

IMMDevice? GetDefaultAudioEndpoint()
{
    var enumerator = (IMMDeviceEnumerator)new MMDeviceEnumerator();
    if (PrintIfFailed(enumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out var current), "Failed to get default audio endpoint"))
    {
        return null;
    }

    return current;
}

List<IMMDevice> GetDevices()
{
    var deviceList = new List<IMMDevice>();
    var enumerator = (IMMDeviceEnumerator)new MMDeviceEnumerator();
    PrintIfFailed(enumerator.EnumAudioEndpoints(EDataFlow.eRender, DeviceState.Active, out var devices), "Failed to enumerate audio endpoints");

    var hr = devices.GetCount(out var cDevices);
    if (hr < 0 || cDevices == 0)
    {
        PrintIfFailed(hr, "Didn't find any configured multimedia devices");
        return deviceList;
    }

    for (var i = 0; i < cDevices; i++)
    {
        hr = devices.Item((uint)i, out var device);
        if (hr < 0 || device == null)
        {
            PrintIfFailed(hr, $"Failed to get device #{i}");
            continue;
        }

        deviceList.Add(device);
    }

    return deviceList;
}

void ListDevices(string[] args)
{
    var devices = GetDevices();

    var current = GetDefaultAudioEndpoint();
    if (current == null)
    {
        return;
    }

    var hr = current.GetId(out var currentId);
    if (hr < 0)
    {
        PrintIfFailed(hr, "Failed to get default audio endpoint id");
        return;
    }

    var defaultForegroundColor = Console.ForegroundColor;
    Console.WriteLine("Device list:");

    for (var i = 0; i < devices.Count; i++)
    {
        var device = devices[i];
        hr = device.GetId(out var deviceId);
        if (hr < 0)
        {
            PrintIfFailed(hr, $"Failed to read the device id for device #{i}");
            continue;
        }

        hr = device.OpenPropertyStore(MMConstants.STGM_READ, out var properties);
        if (hr < 0)
        {
            PrintIfFailed(hr, $"Failed to open the property store for device #{i}");
            continue;
        }

        hr = properties.GetValue(MMConstants.PKEY_Device_FriendlyName, out var pv);
        if (hr < 0)
        {
            PrintIfFailed(hr, $"Failed to read the name for device #{i}");
            continue;
        }

        hr = device.GetState(out var state);
        if (hr < 0)
        {
            PrintIfFailed(hr, $"Failed to read the device state for device #{i}");
            continue;
        }

        var name = Marshal.PtrToStringUni(pv.pwszVal);

        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.Write($"[{i} Id: {deviceId}, State: {state}]");

        var isCurrent = currentId.Equals(deviceId, StringComparison.OrdinalIgnoreCase);
        Console.ForegroundColor = isCurrent ? ConsoleColor.Red : defaultForegroundColor;
        Console.Write(isCurrent ? " (Current Default)" : "");

        Console.ForegroundColor = defaultForegroundColor;
        Console.WriteLine(" \"{0}\"", name);
    }
}

void PrintUsage(string[] args)
{
    Console.WriteLine("Audio Device Toggle");
    Console.WriteLine("Usage: AudioDeviceToggle.exe [options]");
    Console.WriteLine("Options:");

    foreach (var command in commands)
    {
        Console.WriteLine($"({command.Name}) {string.Join(", ", command.Aliases)}");
    }
}

void CycleDevices(string[] args)
{
    if (args.Length < 3)
    {
        Console.WriteLine("At least two device ids must be specified to cycle through.");
        return;
    }


    var defaultDevice = GetDefaultAudioEndpoint();
    if (defaultDevice == null)
    {
        return;
    }

    if (PrintIfFailed(defaultDevice.GetId(out var defaultDeviceId), "Failed to read default device id"))
    {
        return;
    }

    var indexOfCurrentDefaultDevice = -1;

    for (var i = 1; i < args.Length; i++)
    {
        if (defaultDeviceId.Equals(args[i], StringComparison.OrdinalIgnoreCase))
        {
            indexOfCurrentDefaultDevice = i - 1;
            break;
        }
    }

    var nextDeviceId = args[1 + ((indexOfCurrentDefaultDevice + 1) % (args.Length - 1))];

    Console.WriteLine($"Switching default device from id: {defaultDeviceId} to id: {nextDeviceId}");

    SetDefaultDeviceById(["", nextDeviceId]);
}

void SetDefaultDeviceByName(string[] args)
{
    if (args.Length < 2)
    {
        Console.WriteLine("Device name not specified.");
        return;
    }

    var deviceName = args[1];

    var devices = GetDevices();

    var device = devices.FirstOrDefault(d =>
    {
        if (PrintIfFailed(d.GetId(out var id), "Failed to read device id"))
        {
            return false;
        }

        if (PrintIfFailed(d.OpenPropertyStore(MMConstants.STGM_READ, out var properties), $"Failed to open the property store for device #{id}"))
        {
            return false;
        }

        if (PrintIfFailed(properties.GetValue(MMConstants.PKEY_Device_FriendlyName, out var pv), $"Failed to read the name for device #{id}"))
        {
            return false;
        }

        var name = Marshal.PtrToStringUni(pv.pwszVal);
        return name.Equals(deviceName, StringComparison.OrdinalIgnoreCase);
    });

    if (device == null)
    {
        Console.WriteLine($"Could not find device with name: {deviceName}");
        return;
    }

    if (PrintIfFailed(device.GetId(out var deviceId), "Failed to read device id"))
    {
        return;
    }

    SetDefaultDeviceById(["", deviceId]);
}

void SetDefaultDeviceById(string[] args)
{
    if (args.Length < 2)
    {
        Console.WriteLine("Device id not specified.");
        return;
    }

    var newDeviceId = args[1];

    var policyConfig = (IPolicyConfig)new PolicyConfig();
    var devices = GetDevices();

    var newDevice = devices.FirstOrDefault(d =>
    {
        if (PrintIfFailed(d.GetId(out var id), "Failed to read device id"))
        {
            return false;
        }

        return id.Equals(newDeviceId, StringComparison.OrdinalIgnoreCase);
    });

    if (newDevice == null)
    {
        Console.WriteLine($"Could not find device with id: {newDeviceId}");
        return;
    }

    if (PrintIfFailed(policyConfig.SetDefaultEndpoint(newDeviceId, ERole.eMultimedia), "Failed to set default endpoint for role Multimedia"))
    {
        return;
    }

    Console.WriteLine($"Set default endpoint to id: {newDeviceId} for role {ERole.eMultimedia}");

    /*
    foreach (var role in Enum.GetValues<ERole>())
    {
        var hr = policyConfig.SetDefaultEndpoint(deviceId, role);
        if (PrintIfFailed(hr, $"Failed to set default endpoint for role {role}"))
        {
            return;
        }

        Console.WriteLine($"Set default endpoint for role {role}");
    }*/
}

bool PrintIfFailed(int hResult, string message)
{
    if (hResult >= 0)
    {
        return false;
    }

    Console.WriteLine($"{message} (hr = 0x{hResult:X8})");
    return true;
}

bool MatchesArgument(string arg, CommandLineArg command)
{
    return Array.Exists(command.Aliases, a => $"/{a}".Equals(arg, StringComparison.OrdinalIgnoreCase) ||
                                              $"-{a}".Equals(arg, StringComparison.OrdinalIgnoreCase));
}

internal record CommandLineArg(string Name, string[] Aliases, Action<string[]> Callback)
{
    public string Name = Name;
    public string[] Aliases = Aliases;
    public Action<string[]> Callback = Callback;
}
