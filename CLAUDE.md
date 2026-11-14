# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

Windows-only .NET 10 CLI that lists, sets, and cycles the default audio output device. See README.md for user-facing commands and matching rules.

## Commands

```
dotnet build -c Release                # build
dotnet run -- list                     # run with args (read-only command)
dotnet test AudioDeviceToggle.slnx     # runs test projects in the solution (none exist yet)
dotnet publish -c Release              # Native AOT exe (~2 MB) -> bin\Release\net10.0-windows\win-x64\publish\
```

Publishing is Native AOT, which needs the MSVC linker (Visual Studio "Desktop development with C++" workload). The ILCompiler targets invoke `vswhere.exe` by name, so if publish fails with "'vswhere.exe' is not recognized", put `%ProgramFiles(x86)%\Microsoft Visual Studio\Installer` on PATH. `dotnet build`/`dotnet run` don't need this.

There are no test projects and no linter configured. Any command other than `list`/`help` changes the machine's real default audio device and plays a sound — don't run them casually to "verify" changes.

CI (`.github/workflows/build.yml`, windows-latest) restores, runs `dotnet test` on the solution, publishes, and uploads only `AudioDeviceToggle.exe` as the `AudioDeviceToggle-win-x64` artifact. New test projects must be added to `AudioDeviceToggle.slnx` to run in CI.

Releases: pushing a `v*` tag (e.g. `git tag v1.2.0` then `git push origin v1.2.0`) stamps the exe with that version and creates a GitHub release with the exe attached and generated notes. Tags containing `-` (`v1.2.0-beta.1`) become pre-releases. Branch/PR builds never create releases.

## Architecture

- `Program.cs` — top-level statements. Commands are a `List<CommandLineArg>` (name, aliases, args hint, callback); `help` is appended after the initializer because `PrintUsage` reads the list. Device lookup goes through a `DeviceKey` (which field to match, whether partial substring matching is allowed: names yes, ids no). Cycling resolves candidates to device ids before comparing, so partial names still identify the current device.
- `AudioDevices.cs` — the only place that uses the Core Audio COM interfaces; callers only see plain `AudioDevice(Id, Name)` records. `SetDefault` sets the eConsole and eMultimedia roles only (eCommunications deliberately untouched).
- `CoreAudio.cs` / `IPolicyConfig.cs` — hand-written COM bindings (`IMMDeviceEnumerator` & co., and the undocumented `IPolicyConfig`). They must use `[GeneratedComInterface]` + `Com.CreateInstance<T>`, not `[ComImport]`/`new` on a coclass: built-in COM interop doesn't exist under Native AOT and fails only at runtime. COM dispatches by vtable slot: never reorder or remove methods before the last one called, even unused ones.
- Any new dependency must be AOT-compatible — check publish output for IL2xxx/IL3xxx warnings, but also test the published exe: P/Invoke code that passes classes and relies on the JIT pinning them (e.g. NAudio) breaks under AOT with no warning.

## Conventions

- Error handling: `Fail()` is `[DoesNotReturn]` — writes to stderr, plays the failure sound, exits 1. One top-level `try/catch` routes unexpected exceptions (e.g. COM errors) through `Fail`; don't add per-command catches.
- Sounds are embedded resources with `LogicalName="%(Filename)%(Extension)"`, loaded by bare filename via `GetManifestResourceStream`. The published exe must stay standalone (no loose files).
- `WavePlayer.cs` plays sounds with winmm `PlaySound` (PCM 8/16-bit only), applying volume by scaling samples. NAudio was removed because its `WaveOutEvent` fails silently under AOT — the error only surfaces via `PlaybackStopped`. Lesson: "no exception" is not proof of playback; verify by timing (`bogus` command plays the ~2.1 s failure sound).
