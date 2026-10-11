using System.Diagnostics.CodeAnalysis;
using AudioDeviceToggle;
using NAudio.Wave;

const float SoundVolume = 0.15f;
const string SuccessSound = "Long_External15_343.wav";
const string FailureSound = "homeless_exact1.wav";

var byId = new DeviceKey("id", d => d.Id, AllowPartialMatch: false);
// Partial matching lets "HyperX" survive Windows renaming a device (e.g. adding a "2- " prefix).
var byName = new DeviceKey("name", d => d.Name, AllowPartialMatch: true);

var commands = new List<CommandLineArg>
{
    new("list", ["list", "l"], "", _ => ListDevices()),
    new("set device by id", ["id", "i"], "<id>", a => HandleSetDefault(a, byId)),
    new("set device by name", ["name", "n"], "<name>", a => HandleSetDefault(a, byName)),
    new("cycle devices by id", ["ci"], "<id> <id> [...]", a => CycleDevices(a, byId)),
    new("cycle devices by name", ["c"], "<name> <name> [...]", a => CycleDevices(a, byName)),
};

// Added separately: PrintUsage reads `commands`, which can't be referenced inside its own initializer.
commands.Add(new CommandLineArg("help", ["?", "h", "help"], "", _ => PrintUsage()));

// Single catch point: anything unexpected (e.g. COM errors) still reports and plays the failure sound.
try
{
    if (args.Length == 0)
    {
        PrintUsage();
        Console.WriteLine();
        ListDevices();
    }
    else
    {
        var matched = commands.FirstOrDefault(c => MatchesArgument(args[0], c));
        if (matched == null)
        {
            PrintUsage();
            Fail($"Unknown command: {args[0]}");
        }

        matched.Callback(args[1..]);
    }
}
catch (Exception ex)
{
    Fail(ex.Message);
}

return;

void ListDevices()
{
    var devices = AudioDevices.GetRenderDevices();
    var currentId = AudioDevices.GetDefault()?.Id;

    Console.WriteLine("Device list:");

    for (var i = 0; i < devices.Count; i++)
    {
        var device = devices[i];
        var isCurrent = device.Id.Equals(currentId, StringComparison.OrdinalIgnoreCase);

        WriteColored($"{i + 1}.", ConsoleColor.Cyan);
        Console.Write($" \"{device.Name}\"");
        if (isCurrent)
        {
            WriteColored(" (Current Default)", ConsoleColor.Red);
        }

        Console.WriteLine();
    }
}

void PrintUsage()
{
    Console.WriteLine("Usage:");

    foreach (var command in commands)
    {
        WriteColored($"{command.Name} ({string.Join(", ", command.Aliases)})", ConsoleColor.Yellow);
        Console.WriteLine(command.Args.Length > 0 ? $" {command.Args}" : string.Empty);
    }
}

void HandleSetDefault(string[] commandArgs, DeviceKey key)
{
    if (commandArgs.Length < 1)
    {
        Fail($"Device {key.Kind} not specified.");
    }

    SetDefaultDeviceBy(key, commandArgs[0]);
}

void CycleDevices(string[] candidates, DeviceKey key)
{
    if (candidates.Length < 2)
    {
        Fail($"At least two device {key.Kind}s must be specified to cycle through.");
    }

    // Compare resolved device ids, not raw strings, so partial names still identify the current device.
    var devices = AudioDevices.GetRenderDevices();
    var current = AudioDevices.GetDefault();
    var currentIndex = current == null ? -1 : Array.FindIndex(candidates, c => FindDevice(devices, key, c)?.Id == current.Id);
    var next = candidates[(currentIndex + 1) % candidates.Length];

    Console.WriteLine("Switching device");
    Console.Write("from: ");
    WriteColored(current?.Name ?? "(none)", ConsoleColor.Magenta, newline: true);
    Console.Write("to: ");
    WriteColored(next, ConsoleColor.Magenta, newline: true);

    SetDefaultDeviceBy(key, next);
}

// Exact match first; if the key allows it, falls back to a substring match that must hit exactly one device.
AudioDevice? FindDevice(IReadOnlyList<AudioDevice> devices, DeviceKey key, string value)
{
    var exact = devices.FirstOrDefault(d => key.Select(d).Equals(value, StringComparison.OrdinalIgnoreCase));
    if (exact != null || !key.AllowPartialMatch)
    {
        return exact;
    }

    var partial = devices.Where(d => key.Select(d).Contains(value, StringComparison.OrdinalIgnoreCase)).ToList();
    if (partial.Count > 1)
    {
        Fail($"\"{value}\" matches multiple devices: {string.Join(", ", partial.Select(d => $"\"{d.Name}\""))}");
    }

    return partial.FirstOrDefault();
}

void SetDefaultDeviceBy(DeviceKey key, string value)
{
    var device = FindDevice(AudioDevices.GetRenderDevices(), key, value);
    if (device == null)
    {
        ListDevices();
        Fail($"Could not find device with {key.Kind}: {value}");
    }

    var hr = AudioDevices.SetDefault(device.Id);
    if (hr.Failed)
    {
        Fail($"Failed to set default endpoint (hr = 0x{(uint)hr:X8})");
    }

    WriteColored("Default endpoint changed to ", ConsoleColor.Yellow);
    WriteColored($"\"{device.Name}\"", ConsoleColor.White, newline: true);
    Console.WriteLine();

    ListDevices();

    Console.WriteLine("Playing Sound!");
    PlayEmbeddedSound(SuccessSound);

    Console.WriteLine("Done!");
}

// Aliases may be prefixed with '/', '-' or '--'.
bool MatchesArgument(string arg, CommandLineArg command)
{
    return command.Aliases.Contains(arg.TrimStart('/', '-'), StringComparer.OrdinalIgnoreCase);
}

void WriteColored(string text, ConsoleColor color, bool newline = false, TextWriter? writer = null)
{
    writer ??= Console.Out;
    var previous = Console.ForegroundColor;
    Console.ForegroundColor = color;
    writer.Write(newline ? text + Environment.NewLine : text);
    Console.ForegroundColor = previous;
}

/*
 * Prints the error to stderr, plays the failure sound and exits with code 1.
 */
[DoesNotReturn]
void Fail(string message)
{
    WriteColored("Error: ", ConsoleColor.Red, writer: Console.Error);
    WriteColored(message, ConsoleColor.White, newline: true, writer: Console.Error);

    try
    {
        Console.Error.WriteLine("Playing Failure Sound!");
        PlayEmbeddedSound(FailureSound);
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"Failed to play failure sound: {ex.Message}");
    }

    Environment.Exit(1);
}

/*
 * Plays a WAV embedded in the assembly (see EmbeddedResource in the .csproj) and blocks until playback finishes.
 */
void PlayEmbeddedSound(string resourceName)
{
    using var stream = typeof(CommandLineArg).Assembly.GetManifestResourceStream(resourceName)
                       ?? throw new FileNotFoundException($"Embedded sound not found: {resourceName}");
    using var reader = new WaveFileReader(stream);
    using var outputDevice = new WaveOutEvent();
    outputDevice.Volume = SoundVolume;
    outputDevice.Init(reader);
    outputDevice.Play();

    while (outputDevice.PlaybackState == PlaybackState.Playing)
    {
        Thread.Sleep(100);
    }
}

internal record CommandLineArg(string Name, string[] Aliases, string Args, Action<string[]> Callback);

// Which AudioDevice field a command matches its arguments against.
internal record DeviceKey(string Kind, Func<AudioDevice, string> Select, bool AllowPartialMatch);
