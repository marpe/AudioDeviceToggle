# AudioDeviceToggle

A Windows CLI tool to list, set, and cycle default audio output devices.

## Requirements

- Windows
- .NET 10.0

## Usage

```
AudioDeviceToggle <command> [args...]
```

### Commands

| Command | Aliases | Description |
|---------|---------|-------------|
| list | `list`, `l` | List all active audio output devices |
| set device by id | `id`, `i` | Set default device by device ID |
| set device by name | `name`, `n` | Set default device by friendly name |
| cycle devices by id | `ci` | Cycle through devices by ID |
| cycle devices by name | `c` | Cycle through devices by name |
| help | `?`, `h`, `help` | Show usage |

Aliases can be prefixed with `/`, `-`, or `--`.

### Examples

List available devices:
```
AudioDeviceToggle list
```

Set default device by name:
```
AudioDeviceToggle name "Speakers (Focusrite USB Audio)"
```

Cycle between two devices by name (switches to the next one each time):
```
AudioDeviceToggle c "Speakers (HyperX Cloud Flight Wireless Headset)" "Speakers (Focusrite USB Audio)"
```

Cycle between devices by ID:
```
AudioDeviceToggle ci "{0.0.0.00000000}.{d0d43511-...}" "{0.0.0.00000000}.{0dc6ae6b-...}"
```

## How it works

Uses the undocumented Windows COM interface `IPolicyConfig` to set the default audio render endpoint, and the Windows Core Audio API (`IMMDeviceEnumerator`) to enumerate devices. After switching, a short confirmation sound is played on the new device.

## Dependencies

- [NAudio](https://github.com/naudio/NAudio) — audio playback
- [Vanara.PInvoke.CoreAudio](https://github.com/dahall/Vanara) — Windows Core Audio API bindings
