using System.Buffers.Binary;
using System.Runtime.InteropServices;

namespace AudioDeviceToggle;

internal static partial class WavePlayer
{
    private const uint SND_SYNC = 0x0000;
    private const uint SND_NODEFAULT = 0x0002;
    private const uint SND_MEMORY = 0x0004;
    private const ushort WAVE_FORMAT_PCM = 1;

    public static void Play(byte[] wav, float volume)
    {
        ScaleVolume(wav, volume);
        if (!PlaySound(wav, 0, SND_MEMORY | SND_SYNC | SND_NODEFAULT))
        {
            throw new InvalidOperationException("PlaySound failed.");
        }
    }

    private static void ScaleVolume(Span<byte> wav, float volume)
    {
        var data = FindPcmData(wav, out var bitsPerSample);
        switch (bitsPerSample)
        {
            case 8: // unsigned, centred on 128
                foreach (ref var sample in data)
                {
                    sample = (byte)(128 + (sample - 128) * volume);
                }

                break;
            case 16:
                for (var i = 0; i + 1 < data.Length; i += 2)
                {
                    var sample = BinaryPrimitives.ReadInt16LittleEndian(data[i..]);
                    BinaryPrimitives.WriteInt16LittleEndian(data[i..], (short)(sample * volume));
                }

                break;
            default:
                throw new NotSupportedException($"Unsupported WAV bit depth: {bitsPerSample}");
        }
    }

    private static Span<byte> FindPcmData(Span<byte> wav, out int bitsPerSample)
    {
        if (wav.Length < 12 || !wav[..4].SequenceEqual("RIFF"u8) || !wav[8..12].SequenceEqual("WAVE"u8))
        {
            throw new InvalidDataException("Not a WAV file.");
        }

        bitsPerSample = 0;
        var offset = 12;
        while (offset + 8 <= wav.Length)
        {
            var id = wav.Slice(offset, 4);
            var size = BinaryPrimitives.ReadInt32LittleEndian(wav[(offset + 4)..]);
            var body = wav.Slice(offset + 8, Math.Min(size, wav.Length - offset - 8));

            if (id.SequenceEqual("fmt "u8))
            {
                if (BinaryPrimitives.ReadUInt16LittleEndian(body) != WAVE_FORMAT_PCM)
                {
                    throw new NotSupportedException("Only PCM WAV files are supported.");
                }

                bitsPerSample = BinaryPrimitives.ReadUInt16LittleEndian(body[14..]);
            }
            else if (id.SequenceEqual("data"u8))
            {
                if (bitsPerSample == 0)
                {
                    throw new InvalidDataException("WAV data chunk precedes fmt chunk.");
                }

                return body;
            }

            // Chunks are padded to an even size.
            offset += 8 + size + (size & 1);
        }

        throw new InvalidDataException("WAV has no data chunk.");
    }

    [LibraryImport("winmm.dll", EntryPoint = "PlaySoundW")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool PlaySound(byte[] sound, nint module, uint flags);
}
