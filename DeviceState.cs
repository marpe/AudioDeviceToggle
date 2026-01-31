namespace AudioDeviceToggle;

[Flags]
public enum DeviceState
{
    Active = 1,
    Disabled = 2,
    NotPresent = 4,
    Unplugged = 8,
    All = Unplugged | NotPresent | Disabled | Active, // 0x0000000F
}
