# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

Windows-only .NET 10 CLI that lists, sets, and cycles the default audio output device. See README.md for user-facing commands and matching rules.

## Commands

```
dotnet build -c Release                # build
dotnet run -- list                     # run with args (read-only command)
dotnet test AudioDeviceToggle.slnx     # runs test projects in the solution (none exist yet)
dotnet publish -c Release              # single self-contained exe -> bin\Release\net10.0-windows\win-x64\publish\
```

There are no test projects and no linter configured. Any command other than `list`/`help` changes the machine's real default audio device and plays a sound — don't run them casually to "verify" changes.

CI (`.github/workflows/build.yml`, windows-latest) restores, runs `dotnet test` on the solution, publishes, and uploads only `AudioDeviceToggle.exe` as the `AudioDeviceToggle-win-x64` artifact. New test projects must be added to `AudioDeviceToggle.slnx` to run in CI.

Releases: pushing a `v*` tag (e.g. `git tag v1.2.0` then `git push origin v1.2.0`) stamps the exe with that version and creates a GitHub release with the exe attached and generated notes. Tags containing `-` (`v1.2.0-beta.1`) become pre-releases. Branch/PR builds never create releases.

## Architecture

- `Program.cs` — top-level statements. Commands are a `List<CommandLineArg>` (name, aliases, args hint, callback); `help` is appended after the initializer because `PrintUsage` reads the list. Device lookup goes through a `DeviceKey` (which field to match, whether partial substring matching is allowed: names yes, ids no). Cycling resolves candidates to device ids before comparing, so partial names still identify the current device.
- `AudioDevices.cs` — the only place that touches Core Audio COM (via Vanara). Every COM object is wrapped in `ComReleaserFactory.Create(...)` and released before returning; callers only see plain `AudioDevice(Id, Name)` records. `SetDefault` sets the eConsole and eMultimedia roles only (eCommunications deliberately untouched).
- `IPolicyConfig.cs` — undocumented `IPolicyConfig` COM interface. COM dispatches by vtable slot: never reorder or remove methods, even unused ones.

## Conventions

- Error handling: `Fail()` is `[DoesNotReturn]` — writes to stderr, plays the failure sound, exits 1. One top-level `try/catch` routes unexpected exceptions (e.g. COM errors) through `Fail`; don't add per-command catches.
- Sounds are embedded resources with `LogicalName="%(Filename)%(Extension)"`, loaded by bare filename via `GetManifestResourceStream`. The published exe must stay standalone (no loose files).
- Reference `NAudio.WinMM`, not the `NAudio` meta-package — on `-windows` TFMs the meta-package drags in WinForms (~14 MB).
