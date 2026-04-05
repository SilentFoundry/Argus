using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;

namespace FileOrganizer;

internal static class PdfNamingService
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

    public static string ResolveBaseName(FileInfo file, PdfRenameMode pdfRenameMode, string invalidCharacterReplacement)
    {
        var originalName = Path.GetFileNameWithoutExtension(file.Name);
        if (!string.Equals(file.Extension, ".pdf", StringComparison.OrdinalIgnoreCase) ||
            pdfRenameMode == PdfRenameMode.Disabled)
        {
            return originalName;
        }

        var metadata = TryGetPdfNamingData(file, pdfRenameMode, invalidCharacterReplacement);
        if (metadata is null)
        {
            return originalName;
        }

        return metadata.FormattedName;
    }

    private static PdfNamingData? TryGetPdfNamingData(FileInfo file, PdfRenameMode pdfRenameMode, string invalidCharacterReplacement)
    {
        try
        {
            using var document = PdfDocument.Open(file.FullName);

            var documentTitle = NormalizeCandidate(document.Information?.Title);
            var firstPagesText = GetFirstPagesText(document);
            var contentHeading = ExtractHeadingFromContent(firstPagesText);

            var selectedTitle = pdfRenameMode switch
            {
                PdfRenameMode.TitleOnly => ChooseTitleCandidate(documentTitle, null, file),
                PdfRenameMode.ContentHeadingOnly => ChooseTitleCandidate(null, contentHeading, file),
                PdfRenameMode.TitleThenContent => ChooseTitleCandidate(documentTitle, contentHeading, file),
                _ => ChooseTitleCandidate(null, null, file)
            };

            var relevantDate = ExtractDateFromContent(firstPagesText)
                ?? document.Information?.GetCreatedDateTimeOffset()?.Date
                ?? file.CreationTime.Date;

            var shortTitle = BuildShortTitle(selectedTitle, invalidCharacterReplacement);
            return new PdfNamingData($"{relevantDate:yyyy-MM-dd}_{shortTitle}");
        }
        catch
        {
            return null;
        }
    }

    private static string GetFirstPagesText(PdfDocument document)
    {
        var pagesToInspect = Math.Min(document.NumberOfPages, 2);
        var parts = Enumerable.Range(1, pagesToInspect)
            .Select(pageNumber => document.GetPage(pageNumber).Text ?? string.Empty)
            .Where(text => !string.IsNullOrWhiteSpace(text))
            .ToList();

        return string.Join(Environment.NewLine, parts);
    }

    private static string ChooseTitleCandidate(string? documentTitle, string? contentHeading, FileInfo file)
    {
        if (IsUsefulCandidate(documentTitle))
        {
            return documentTitle!;
        }

        if (IsUsefulCandidate(contentHeading))
        {
            return contentHeading!;
        }

        return Path.GetFileNameWithoutExtension(file.Name);
    }

    private static DateTime? ExtractDateFromContent(string text)
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

    private static string? ExtractHeadingFromContent(string text)
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

    private static bool IsUsefulCandidate(string? candidate)
    {
        if (string.IsNullOrWhiteSpace(candidate))
        {
            return false;
        }

        if (candidate.Length < 6 || candidate.Length > 140)
        {
            return false;
        }

        if (SkipHeadings.Contains(candidate, StringComparer.OrdinalIgnoreCase))
        {
            return false;
        }

        var letterCount = candidate.Count(char.IsLetter);
        if (letterCount < 4)
        {
            return false;
        }

        var wordCount = candidate.Split(' ', StringSplitOptions.RemoveEmptyEntries).Length;
        if (wordCount < 2 || wordCount > 18)
        {
            return false;
        }

        if (candidate.StartsWith("page ", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return true;
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

    private sealed record PdfNamingData(string FormattedName);
}
