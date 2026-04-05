using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace FileOrganizer;

internal static class MediaMetadataService
{
    private static readonly int[] ExifDatePropertyIds =
    {
        0x9003,
        0x9004,
        0x0132
    };

    public static string ResolveImageBaseName(FileInfo file, string invalidCharacterReplacement)
    {
        var originalName = Path.GetFileNameWithoutExtension(file.Name);

        try
        {
            using var image = Image.FromFile(file.FullName);
            var dateTaken = TryGetImageDateTaken(image) ?? file.CreationTime;
            var safeName = SanitizeBaseName(originalName, invalidCharacterReplacement);
            return $"{dateTaken:yyyy-MM-dd}_{safeName}";
        }
        catch
        {
            return originalName;
        }
    }

    private static DateTime? TryGetImageDateTaken(Image image)
    {
        foreach (var propertyId in ExifDatePropertyIds)
        {
            var property = image.PropertyItems.FirstOrDefault(item => item.Id == propertyId);
            var bytes = property?.Value;
            if (bytes is not { Length: > 0 })
            {
                continue;
            }

            try
            {
                var raw = Encoding.ASCII.GetString(bytes).Trim('\0', ' ', '\t', '\r', '\n');
                if (string.IsNullOrWhiteSpace(raw))
                {
                    continue;
                }

                if (DateTime.TryParseExact(raw, "yyyy:MM:dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out var exact))
                {
                    return exact;
                }

                if (DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var parsed))
                {
                    return parsed;
                }
            }
            catch
            {
                // Ignore malformed metadata and continue.
            }
        }

        return null;
    }

    private static string SanitizeBaseName(string name, string replacement)
    {
        var result = name;
        foreach (var invalidChar in Path.GetInvalidFileNameChars())
        {
            result = result.Replace(invalidChar.ToString(), replacement, StringComparison.Ordinal);
        }

        result = result.Trim().TrimEnd('.');
        return string.IsNullOrWhiteSpace(result) ? "image" : result;
    }
}

