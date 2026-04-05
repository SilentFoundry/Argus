using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace FileOrganizer;

internal static class MetadataNamingService
{
    private static readonly string[] SkipHeadings =
    {
        "abstract",
        "contents",
        "table of contents",
        "keywords",
        "references",
        "appendix"
    };

    private static readonly Regex[] DatePatterns =
    {
        new(@"\b\d{4}-\d{1,2}-\d{1,2}\b", RegexOptions.IgnoreCase | RegexOptions.Compiled),
        new(@"\b\d{1,2}/\d{1,2}/\d{2,4}\b", RegexOptions.IgnoreCase | RegexOptions.Compiled),
        new(@"\b(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\s+\d{1,2},\s+\d{4}\b", RegexOptions.IgnoreCase | RegexOptions.Compiled),
        new(@"\b\d{1,2}\s+(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\s+\d{4}\b", RegexOptions.IgnoreCase | RegexOptions.Compiled)
    };

    public static string ResolveBaseName(FileInfo file, PdfRenameMode renameMode, string invalidCharacterReplacement)
    {
        var originalName = Path.GetFileNameWithoutExtension(file.Name);
        if (renameMode == PdfRenameMode.Disabled)
        {
            return originalName;
        }

        var extension = file.Extension.ToLowerInvariant();
        var resolved = extension switch
        {
            ".pdf" => PdfNamingService.ResolveBaseName(file, renameMode, invalidCharacterReplacement),
            ".docx" => TryResolveDocxName(file, renameMode, invalidCharacterReplacement),
            ".pptx" => TryResolvePptxName(file, renameMode, invalidCharacterReplacement),
            ".xlsx" => TryResolveXlsxName(file, renameMode, invalidCharacterReplacement),
            ".epub" => TryResolveEpubName(file, renameMode, invalidCharacterReplacement),
            ".mp3" => TryResolveMp3Name(file, renameMode, invalidCharacterReplacement),
            ".m4a" or ".mp4" or ".m4v" or ".mov" => TryResolveMp4LikeName(file, renameMode, invalidCharacterReplacement),
            _ => null
        };

        return string.IsNullOrWhiteSpace(resolved) ? originalName : resolved!;
    }

    private static string? TryResolveDocxName(FileInfo file, PdfRenameMode renameMode, string invalidCharacterReplacement)
    {
        try
        {
            using var archive = ZipFile.OpenRead(file.FullName);
            var core = ReadOpenXmlCoreProperties(archive);
            var content = ReadXmlEntryText(archive, "word/document.xml", maxValues: 120);
            return BuildFormattedName(file, renameMode, invalidCharacterReplacement, core.Title, content, core.Date);
        }
        catch
        {
            return null;
        }
    }

    private static string? TryResolvePptxName(FileInfo file, PdfRenameMode renameMode, string invalidCharacterReplacement)
    {
        try
        {
            using var archive = ZipFile.OpenRead(file.FullName);
            var core = ReadOpenXmlCoreProperties(archive);
            var slideText = archive.Entries
                .Where(entry => entry.FullName.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase) &&
                                entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                .OrderBy(entry => entry.FullName, StringComparer.OrdinalIgnoreCase)
                .Take(3)
                .Select(entry => string.Join(Environment.NewLine, ReadXmlEntryValues(entry, maxValues: 60)))
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToList();

            return BuildFormattedName(file, renameMode, invalidCharacterReplacement, core.Title, string.Join(Environment.NewLine, slideText), core.Date);
        }
        catch
        {
            return null;
        }
    }

    private static string? TryResolveXlsxName(FileInfo file, PdfRenameMode renameMode, string invalidCharacterReplacement)
    {
        try
        {
            using var archive = ZipFile.OpenRead(file.FullName);
            var core = ReadOpenXmlCoreProperties(archive);
            var workbookText = ReadWorkbookText(archive);
            return BuildFormattedName(file, renameMode, invalidCharacterReplacement, core.Title, workbookText, core.Date);
        }
        catch
        {
            return null;
        }
    }

    private static string? TryResolveEpubName(FileInfo file, PdfRenameMode renameMode, string invalidCharacterReplacement)
    {
        try
        {
            using var archive = ZipFile.OpenRead(file.FullName);
            var containerEntry = FindZipEntry(archive, "META-INF/container.xml");
            if (containerEntry is null)
            {
                return null;
            }

            var containerDocument = LoadXml(containerEntry);
            var packagePath = containerDocument
                .Descendants()
                .FirstOrDefault(element => element.Name.LocalName == "rootfile")
                ?.Attribute("full-path")
                ?.Value;

            if (string.IsNullOrWhiteSpace(packagePath))
            {
                return null;
            }

            var packageEntry = FindZipEntry(archive, packagePath);
            if (packageEntry is null)
            {
                return null;
            }

            var packageDocument = LoadXml(packageEntry);
            var title = NormalizeCandidate(packageDocument.Descendants().FirstOrDefault(element => element.Name.LocalName == "title")?.Value);
            var metadataDate = TryParseMetadataDate(packageDocument.Descendants().FirstOrDefault(element => element.Name.LocalName == "date")?.Value);
            var content = ReadEpubContent(archive, packagePath, packageDocument);
            return BuildFormattedName(file, renameMode, invalidCharacterReplacement, title, content, metadataDate);
        }
        catch
        {
            return null;
        }
    }

    private static string? TryResolveMp3Name(FileInfo file, PdfRenameMode renameMode, string invalidCharacterReplacement)
    {
        try
        {
            var metadata = TryReadId3V2(file) ?? TryReadId3V1(file);
            if (metadata is null)
            {
                return null;
            }

            var title = BuildAudioDisplayTitle(metadata.Artist, metadata.Title, metadata.Album);
            var content = string.Join(Environment.NewLine, new[] { metadata.Title, metadata.Artist, metadata.Album }.Where(value => !string.IsNullOrWhiteSpace(value)));
            return BuildFormattedName(file, renameMode, invalidCharacterReplacement, title, content, metadata.Date);
        }
        catch
        {
            return null;
        }
    }

    private static string? TryResolveMp4LikeName(FileInfo file, PdfRenameMode renameMode, string invalidCharacterReplacement)
    {
        try
        {
            var metadata = TryReadMp4Metadata(file);
            if (metadata is null)
            {
                return null;
            }

            var isAudioLike = string.Equals(file.Extension, ".m4a", StringComparison.OrdinalIgnoreCase);
            var title = isAudioLike
                ? BuildAudioDisplayTitle(metadata.Artist, metadata.Title, null)
                : NormalizeCandidate(metadata.Title);
            var content = string.Join(Environment.NewLine, new[] { metadata.Title, metadata.Artist, metadata.Description }.Where(value => !string.IsNullOrWhiteSpace(value)));
            return BuildFormattedName(file, renameMode, invalidCharacterReplacement, title, content, metadata.Date);
        }
        catch
        {
            return null;
        }
    }

    private static string BuildFormattedName(FileInfo file, PdfRenameMode renameMode, string invalidCharacterReplacement, string? metadataTitle, string? contentText, DateTime? metadataDate)
    {
        var originalName = Path.GetFileNameWithoutExtension(file.Name);
        var contentHeading = ExtractHeadingFromText(contentText);
        var selectedTitle = renameMode switch
        {
            PdfRenameMode.TitleOnly => ChooseStrictTitle(metadataTitle, originalName),
            PdfRenameMode.ContentHeadingOnly => ChooseStrictTitle(contentHeading, originalName),
            _ => ChoosePreferredTitle(metadataTitle, contentHeading, originalName)
        };

        var relevantDate = ExtractDateFromText(contentText)
            ?? metadataDate
            ?? file.CreationTime.Date;

        var shortTitle = BuildShortTitle(selectedTitle, invalidCharacterReplacement);
        return $"{relevantDate:yyyy-MM-dd}_{shortTitle}";
    }

    private static string ReadWorkbookText(ZipArchive archive)
    {
        var pieces = new List<string>();
        var workbookEntry = FindZipEntry(archive, "xl/workbook.xml");
        if (workbookEntry is not null)
        {
            var workbookDocument = LoadXml(workbookEntry);
            pieces.AddRange(workbookDocument
                .Descendants()
                .Where(element => element.Name.LocalName == "sheet")
                .Select(element => NormalizeCandidate(element.Attribute("name")?.Value))
                .Where(value => !string.IsNullOrWhiteSpace(value))!
                .Take(12)!);
        }

        var sharedStringsEntry = FindZipEntry(archive, "xl/sharedStrings.xml");
        if (sharedStringsEntry is not null)
        {
            pieces.AddRange(ReadXmlEntryValues(sharedStringsEntry, maxValues: 30));
        }

        return string.Join(Environment.NewLine, pieces);
    }

    private static string ReadEpubContent(ZipArchive archive, string packagePath, XDocument packageDocument)
    {
        var manifest = packageDocument
            .Descendants()
            .Where(element => element.Name.LocalName == "item")
            .Select(element => new
            {
                Id = element.Attribute("id")?.Value,
                Href = element.Attribute("href")?.Value
            })
            .Where(item => !string.IsNullOrWhiteSpace(item.Id) && !string.IsNullOrWhiteSpace(item.Href))
            .ToDictionary(item => item.Id!, item => item.Href!, StringComparer.OrdinalIgnoreCase);

        var contentRoot = GetRelativeDirectory(packagePath);
        foreach (var itemRef in packageDocument.Descendants().Where(element => element.Name.LocalName == "itemref").Take(3))
        {
            var idRef = itemRef.Attribute("idref")?.Value;
            if (string.IsNullOrWhiteSpace(idRef) || !manifest.TryGetValue(idRef, out var href))
            {
                continue;
            }

            var contentPath = CombineArchivePath(contentRoot, href);
            var entry = FindZipEntry(archive, contentPath);
            if (entry is null)
            {
                continue;
            }

            try
            {
                var document = LoadXml(entry);
                var values = document
                    .Descendants()
                    .Where(element => element.Name.LocalName is "title" or "h1" or "h2" or "p")
                    .Select(element => NormalizeCandidate(element.Value))
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Take(24)
                    .ToList();

                if (values.Count > 0)
                {
                    return string.Join(Environment.NewLine, values);
                }
            }
            catch
            {
                using var reader = new StreamReader(entry.Open(), Encoding.UTF8, true);
                var raw = reader.ReadToEnd();
                var stripped = Regex.Replace(raw, "<[^>]+>", " ");
                var normalized = NormalizeCandidate(stripped);
                if (!string.IsNullOrWhiteSpace(normalized))
                {
                    return normalized;
                }
            }
        }

        return string.Empty;
    }

    private static OpenXmlCoreProperties ReadOpenXmlCoreProperties(ZipArchive archive)
    {
        var entry = FindZipEntry(archive, "docProps/core.xml");
        if (entry is null)
        {
            return new OpenXmlCoreProperties(null, null);
        }

        var document = LoadXml(entry);
        var title = NormalizeCandidate(document.Descendants().FirstOrDefault(element => element.Name.LocalName == "title")?.Value);
        var created = document.Descendants().FirstOrDefault(element => element.Name.LocalName is "created" or "modified")?.Value;
        return new OpenXmlCoreProperties(title, TryParseMetadataDate(created));
    }

    private static string ReadXmlEntryText(ZipArchive archive, string path, int maxValues)
    {
        var entry = FindZipEntry(archive, path);
        return entry is null ? string.Empty : string.Join(Environment.NewLine, ReadXmlEntryValues(entry, maxValues));
    }

    private static IReadOnlyList<string> ReadXmlEntryValues(ZipArchiveEntry entry, int maxValues)
    {
        var document = LoadXml(entry);
        return document.Descendants()
            .Where(element => !string.IsNullOrWhiteSpace(element.Value))
            .Select(element => NormalizeCandidate(element.Value))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(maxValues)
            .ToList()!;
    }

    private static XDocument LoadXml(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    private static ZipArchiveEntry? FindZipEntry(ZipArchive archive, string path)
    {
        var normalizedPath = path.Replace('\\', '/');
        return archive.Entries.FirstOrDefault(entry => string.Equals(entry.FullName, normalizedPath, StringComparison.OrdinalIgnoreCase));
    }

    private static string GetRelativeDirectory(string path)
    {
        var normalizedPath = path.Replace('\\', '/');
        var lastSlash = normalizedPath.LastIndexOf('/');
        return lastSlash < 0 ? string.Empty : normalizedPath[..lastSlash];
    }

    private static string CombineArchivePath(string directory, string relativePath)
    {
        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(directory))
        {
            parts.AddRange(directory.Split('/', StringSplitOptions.RemoveEmptyEntries));
        }

        foreach (var part in relativePath.Replace('\\', '/').Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (part == ".")
            {
                continue;
            }

            if (part == "..")
            {
                if (parts.Count > 0)
                {
                    parts.RemoveAt(parts.Count - 1);
                }

                continue;
            }

            parts.Add(part);
        }

        return string.Join("/", parts);
    }

    private static AudioMetadata? TryReadId3V2(FileInfo file)
    {
        using var stream = file.OpenRead();
        if (stream.Length < 10)
        {
            return null;
        }

        var header = ReadRequiredBytes(stream, 10);
        if (header.Length < 10)
        {
            return null;
        }

        if (Encoding.ASCII.GetString(header, 0, 3) != "ID3")
        {
            return null;
        }

        var majorVersion = header[3];
        if (majorVersion is < 3 or > 4)
        {
            return null;
        }

        var tagSize = ReadSyncSafeInteger(header.AsSpan(6, 4));
        var tagEnd = Math.Min(stream.Length, stream.Position + tagSize);
        string? title = null;
        string? artist = null;
        string? album = null;
        DateTime? date = null;

        while (stream.Position + 10 <= tagEnd)
        {
            var frameHeader = ReadRequiredBytes(stream, 10);
            if (frameHeader.Length < 10)
            {
                break;
            }

            if (frameHeader.All(value => value == 0))
            {
                break;
            }

            var frameId = Encoding.ASCII.GetString(frameHeader, 0, 4);
            var frameSize = majorVersion == 4
                ? ReadSyncSafeInteger(frameHeader.AsSpan(4, 4))
                : ReadBigEndianInt32(frameHeader.AsSpan(4, 4));

            if (frameSize <= 0 || stream.Position + frameSize > tagEnd)
            {
                break;
            }

            var frameData = ReadRequiredBytes(stream, frameSize);
            if (frameData.Length < frameSize)
            {
                break;
            }

            if (!frameId.StartsWith('T'))
            {
                continue;
            }

            var text = DecodeId3Text(frameData);
            if (string.IsNullOrWhiteSpace(text))
            {
                continue;
            }

            switch (frameId)
            {
                case "TIT2":
                    title ??= NormalizeCandidate(text);
                    break;
                case "TPE1":
                    artist ??= NormalizeCandidate(text);
                    break;
                case "TALB":
                    album ??= NormalizeCandidate(text);
                    break;
                case "TDRC":
                case "TYER":
                    date ??= TryParseMetadataDate(text) ?? ExtractDateFromText(text);
                    break;
            }
        }

        return string.IsNullOrWhiteSpace(title) && string.IsNullOrWhiteSpace(artist) && string.IsNullOrWhiteSpace(album) && date is null
            ? null
            : new AudioMetadata(title, artist, album, date);
    }

    private static AudioMetadata? TryReadId3V1(FileInfo file)
    {
        using var stream = file.OpenRead();
        if (stream.Length < 128)
        {
            return null;
        }

        stream.Seek(-128, SeekOrigin.End);
        var buffer = ReadRequiredBytes(stream, 128);
        if (buffer.Length < 128)
        {
            return null;
        }

        if (Encoding.ASCII.GetString(buffer, 0, 3) != "TAG")
        {
            return null;
        }

        var title = NormalizeCandidate(Encoding.Latin1.GetString(buffer, 3, 30).Trim('\0', ' '));
        var artist = NormalizeCandidate(Encoding.Latin1.GetString(buffer, 33, 30).Trim('\0', ' '));
        var album = NormalizeCandidate(Encoding.Latin1.GetString(buffer, 63, 30).Trim('\0', ' '));
        var yearText = Encoding.ASCII.GetString(buffer, 93, 4).Trim('\0', ' ');
        var date = TryParseMetadataDate(yearText);

        return string.IsNullOrWhiteSpace(title) && string.IsNullOrWhiteSpace(artist) && string.IsNullOrWhiteSpace(album) && date is null
            ? null
            : new AudioMetadata(title, artist, album, date);
    }

    private static Mp4Metadata? TryReadMp4Metadata(FileInfo file)
    {
        using var stream = file.OpenRead();
        var metadata = new Mp4Metadata();
        ReadMp4Atoms(stream, stream.Length, metadata, depth: 0);
        return string.IsNullOrWhiteSpace(metadata.Title) &&
               string.IsNullOrWhiteSpace(metadata.Artist) &&
               string.IsNullOrWhiteSpace(metadata.Description) &&
               metadata.Date is null
            ? null
            : metadata;
    }

    private static void ReadMp4Atoms(Stream stream, long limit, Mp4Metadata metadata, int depth)
    {
        if (depth > 6)
        {
            return;
        }

        while (stream.Position + 8 <= limit)
        {
            var atomStart = stream.Position;
            var size = ReadUInt32BigEndian(stream);
            var typeBytes = ReadRequiredBytes(stream, 4);
            var type = Encoding.Latin1.GetString(typeBytes);
            long atomSize = size switch
            {
                0 => limit - atomStart,
                1 => ReadInt64BigEndian(stream),
                _ => size
            };

            if (atomSize < 8)
            {
                break;
            }

            var payloadStart = stream.Position;
            var payloadEnd = atomStart + atomSize;
            if (payloadEnd > limit || payloadEnd > stream.Length)
            {
                break;
            }

            switch (type)
            {
                case "moov":
                case "udta":
                case "ilst":
                    ReadMp4Atoms(stream, payloadEnd, metadata, depth + 1);
                    break;
                case "meta":
                    stream.Position = Math.Min(payloadEnd, stream.Position + 4);
                    ReadMp4Atoms(stream, payloadEnd, metadata, depth + 1);
                    break;
                case "\u00a9nam":
                    metadata.Title ??= ReadMp4MetadataValue(stream, payloadEnd);
                    break;
                case "\u00a9day":
                    var dateValue = ReadMp4MetadataValue(stream, payloadEnd);
                    metadata.Date ??= TryParseMetadataDate(dateValue) ?? ExtractDateFromText(dateValue);
                    break;
                case "\u00a9ART":
                case "aART":
                    metadata.Artist ??= ReadMp4MetadataValue(stream, payloadEnd);
                    break;
                case "desc":
                case "ldes":
                    metadata.Description ??= ReadMp4MetadataValue(stream, payloadEnd);
                    break;
            }

            stream.Position = payloadEnd;
            if (stream.Position <= payloadStart)
            {
                break;
            }
        }
    }

    private static string? ReadMp4MetadataValue(Stream stream, long limit)
    {
        while (stream.Position + 8 <= limit)
        {
            var childStart = stream.Position;
            var childSize = ReadUInt32BigEndian(stream);
            var childType = Encoding.Latin1.GetString(ReadRequiredBytes(stream, 4));
            var childEnd = childStart + childSize;
            if (childSize < 8 || childEnd > limit)
            {
                return null;
            }

            if (childType == "data" && stream.Position + 8 <= childEnd)
            {
                stream.Position += 8;
                var payloadLength = (int)(childEnd - stream.Position);
                var payload = ReadRequiredBytes(stream, payloadLength);
                return DecodeMetadataString(payload);
            }

            stream.Position = childEnd;
        }

        return null;
    }

    private static string? DecodeMetadataString(byte[] payload)
    {
        var utf8 = NormalizeCandidate(Encoding.UTF8.GetString(payload).Trim('\0', ' ', '\u0001'));
        if (!string.IsNullOrWhiteSpace(utf8))
        {
            return utf8;
        }

        if (payload.Length >= 2)
        {
            var unicode = NormalizeCandidate(Encoding.Unicode.GetString(payload).Trim('\0', ' '));
            if (!string.IsNullOrWhiteSpace(unicode))
            {
                return unicode;
            }
        }

        return NormalizeCandidate(Encoding.Latin1.GetString(payload).Trim('\0', ' '));
    }

    private static byte[] ReadRequiredBytes(Stream stream, int count)
    {
        var buffer = new byte[count];
        var read = stream.Read(buffer, 0, count);
        return read == count ? buffer : buffer[..read];
    }

    private static int ReadSyncSafeInteger(ReadOnlySpan<byte> bytes)
    {
        return (bytes[0] & 0x7f) << 21 |
               (bytes[1] & 0x7f) << 14 |
               (bytes[2] & 0x7f) << 7 |
               (bytes[3] & 0x7f);
    }

    private static int ReadBigEndianInt32(ReadOnlySpan<byte> bytes)
    {
        return bytes[0] << 24 |
               bytes[1] << 16 |
               bytes[2] << 8 |
               bytes[3];
    }

    private static uint ReadUInt32BigEndian(Stream stream)
    {
        var bytes = ReadRequiredBytes(stream, 4);
        return bytes.Length < 4
            ? 0
            : (uint)(bytes[0] << 24 |
                     bytes[1] << 16 |
                     bytes[2] << 8 |
                     bytes[3]);
    }

    private static long ReadInt64BigEndian(Stream stream)
    {
        var bytes = ReadRequiredBytes(stream, 8);
        if (bytes.Length < 8)
        {
            return 0;
        }

        return ((long)bytes[0] << 56) |
               ((long)bytes[1] << 48) |
               ((long)bytes[2] << 40) |
               ((long)bytes[3] << 32) |
               ((long)bytes[4] << 24) |
               ((long)bytes[5] << 16) |
               ((long)bytes[6] << 8) |
               bytes[7];
    }

    private static string? DecodeId3Text(byte[] frameData)
    {
        if (frameData.Length == 0)
        {
            return null;
        }

        var encodingFlag = frameData[0];
        var content = frameData[1..];
        var text = encodingFlag switch
        {
            0 => Encoding.Latin1.GetString(content),
            1 => Encoding.Unicode.GetString(content),
            2 => Encoding.BigEndianUnicode.GetString(content),
            3 => Encoding.UTF8.GetString(content),
            _ => Encoding.UTF8.GetString(content)
        };

        return NormalizeCandidate(text.Trim('\0', ' '));
    }

    private static string BuildAudioDisplayTitle(string? artist, string? title, string? album)
    {
        var normalizedArtist = NormalizeCandidate(artist);
        var normalizedTitle = NormalizeCandidate(title);
        var normalizedAlbum = NormalizeCandidate(album);

        if (!string.IsNullOrWhiteSpace(normalizedArtist) && !string.IsNullOrWhiteSpace(normalizedTitle))
        {
            return $"{normalizedArtist} {normalizedTitle}";
        }

        if (!string.IsNullOrWhiteSpace(normalizedTitle))
        {
            return normalizedTitle!;
        }

        if (!string.IsNullOrWhiteSpace(normalizedAlbum))
        {
            return normalizedAlbum!;
        }

        return "Audio";
    }

    private static DateTime? TryParseMetadataDate(string? text)
    {
        var candidate = NormalizeCandidate(text);
        if (string.IsNullOrWhiteSpace(candidate))
        {
            return null;
        }

        if (DateTimeOffset.TryParse(candidate, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var parsedOffset))
        {
            return parsedOffset.Date;
        }

        if (DateTime.TryParse(candidate, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var parsed))
        {
            return parsed.Date;
        }

        var yearMatch = Regex.Match(candidate, @"\b(19|20)\d{2}\b");
        return yearMatch.Success && DateTime.TryParseExact(yearMatch.Value, "yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsed)
            ? parsed.Date
            : null;
    }

    private static DateTime? ExtractDateFromText(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        foreach (var pattern in DatePatterns)
        {
            var match = pattern.Match(text);
            if (!match.Success)
            {
                continue;
            }

            if (DateTime.TryParse(match.Value, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var parsed))
            {
                return parsed.Date;
            }

            if (DateTime.TryParse(match.Value, out parsed))
            {
                return parsed.Date;
            }
        }

        return null;
    }

    private static string? ExtractHeadingFromText(string? text)
    {
        var lines = Regex.Split(text ?? string.Empty, @"\r\n|\r|\n")
            .Select(NormalizeCandidate)
            .Where(line => !string.IsNullOrWhiteSpace(line))
            .ToList();

        foreach (var line in lines)
        {
            if (IsUsefulCandidate(line))
            {
                return line;
            }
        }

        return null;
    }

    private static string ChoosePreferredTitle(string? metadataTitle, string? contentHeading, string fallback)
    {
        if (IsUsefulCandidate(metadataTitle))
        {
            return metadataTitle!;
        }

        if (IsUsefulCandidate(contentHeading))
        {
            return contentHeading!;
        }

        return fallback;
    }

    private static string ChooseStrictTitle(string? candidate, string fallback)
    {
        return IsUsefulCandidate(candidate) ? candidate! : fallback;
    }

    private static bool IsUsefulCandidate(string? candidate)
    {
        if (string.IsNullOrWhiteSpace(candidate))
        {
            return false;
        }

        if (candidate.Length < 4 || candidate.Length > 140)
        {
            return false;
        }

        if (SkipHeadings.Contains(candidate, StringComparer.OrdinalIgnoreCase))
        {
            return false;
        }

        var letterCount = candidate.Count(char.IsLetter);
        if (letterCount < 3)
        {
            return false;
        }

        var wordCount = candidate.Split(' ', StringSplitOptions.RemoveEmptyEntries).Length;
        if (wordCount is < 1 or > 18)
        {
            return false;
        }

        return !candidate.StartsWith("page ", StringComparison.OrdinalIgnoreCase);
    }

    private static string? NormalizeCandidate(string? candidate)
    {
        if (string.IsNullOrWhiteSpace(candidate))
        {
            return null;
        }

        var normalized = Regex.Replace(candidate, @"\s+", " ").Trim();
        normalized = normalized.Trim('-', '_', '.', ':', ';', ',', '"', '\'');
        return string.IsNullOrWhiteSpace(normalized) ? null : normalized;
    }

    private static string BuildShortTitle(string title, string invalidCharacterReplacement)
    {
        var cleaned = title;
        foreach (var pattern in DatePatterns)
        {
            cleaned = pattern.Replace(cleaned, " ");
        }

        cleaned = NormalizeCandidate(cleaned) ?? "Document";
        cleaned = Regex.Replace(cleaned, @"[^A-Za-z0-9\s-]", " ");
        cleaned = Regex.Replace(cleaned, @"\s+", " ").Trim();

        var words = cleaned.Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Take(8)
            .ToArray();

        var joined = words.Length == 0 ? "Document" : string.Join("-", words);
        joined = SanitizeFileName(joined, invalidCharacterReplacement);
        return string.IsNullOrWhiteSpace(joined) ? "Document" : joined;
    }

    private static string SanitizeFileName(string name, string replacement)
    {
        var result = name;
        foreach (var invalidChar in Path.GetInvalidFileNameChars())
        {
            result = result.Replace(invalidChar.ToString(), replacement, StringComparison.Ordinal);
        }

        result = Regex.Replace(result, @"\s+", " ").Trim();
        result = Regex.Replace(result, @"-+", "-");
        return result.Trim().TrimEnd('.');
    }

    private sealed record OpenXmlCoreProperties(string? Title, DateTime? Date);
    private sealed record AudioMetadata(string? Title, string? Artist, string? Album, DateTime? Date);
    private sealed class Mp4Metadata
    {
        public string? Title { get; set; }
        public string? Artist { get; set; }
        public string? Description { get; set; }
        public DateTime? Date { get; set; }
    }
}
