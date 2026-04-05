using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace FileOrganizer;

internal static class AppStorage
{
    private const string CurrentDataFolderName = "Argus";
    private const string LegacyDataFolderName = "FileOrganizer";

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    public static AppPaths GetPaths()
    {
        var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var baseDirectory = Path.Combine(localAppData, CurrentDataFolderName);
        var legacyDirectory = Path.Combine(localAppData, LegacyDataFolderName);

        EnsureDataDirectoryMigrated(baseDirectory, legacyDirectory);

        var logsDirectory = Path.Combine(baseDirectory, "logs");
        var reportsDirectory = Path.Combine(baseDirectory, "reports");

        Directory.CreateDirectory(baseDirectory);
        Directory.CreateDirectory(logsDirectory);
        Directory.CreateDirectory(reportsDirectory);

        return new AppPaths(
            baseDirectory,
            Path.Combine(baseDirectory, "settings.json"),
            Path.Combine(baseDirectory, "history.json"),
            logsDirectory,
            reportsDirectory,
            Path.Combine(AppContext.BaseDirectory, "rules.json"));
    }

    public static LoadedRules LoadRules(string? rulesPath = null)
    {
        var paths = GetPaths();
        var resolvedPath = Path.GetFullPath(string.IsNullOrWhiteSpace(rulesPath) ? paths.DefaultRulesPath : rulesPath);
        var defaultConfig = CreateDefaultRules();

        if (!File.Exists(resolvedPath))
        {
            var directory = Path.GetDirectoryName(resolvedPath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.WriteAllText(resolvedPath, JsonSerializer.Serialize(defaultConfig, JsonOptions));
        }

        try
        {
            var loaded = JsonSerializer.Deserialize<RulesConfig>(File.ReadAllText(resolvedPath), JsonOptions) ?? defaultConfig;
            NormalizeRules(loaded);
            return new LoadedRules(resolvedPath, loaded);
        }
        catch (JsonException) when (string.Equals(resolvedPath, paths.DefaultRulesPath, StringComparison.OrdinalIgnoreCase))
        {
            BackupInvalidJsonFile(resolvedPath);
            File.WriteAllText(resolvedPath, JsonSerializer.Serialize(defaultConfig, JsonOptions));
            NormalizeRules(defaultConfig);
            return new LoadedRules(resolvedPath, defaultConfig);
        }
        catch (JsonException ex)
        {
            throw new InvalidOperationException($"Rules file is not valid JSON: {resolvedPath}", ex);
        }
    }

    public static AppSettings LoadSettings()
    {
        var paths = GetPaths();
        return LoadJsonWithFallback(
            paths.SettingsPath,
            () => new AppSettings
            {
                LastRulesPath = paths.DefaultRulesPath,
                WindowLayoutVersion = 1
            },
            settings => string.IsNullOrWhiteSpace(settings.LastRulesPath)
                ? settings with { LastRulesPath = paths.DefaultRulesPath }
                : settings);
    }

    public static void SaveSettings(AppSettings settings)
    {
        var paths = GetPaths();
        File.WriteAllText(paths.SettingsPath, JsonSerializer.Serialize(settings, JsonOptions));
    }

    public static List<RunHistoryEntry> LoadHistory()
    {
        var paths = GetPaths();
        return LoadJsonWithFallback(paths.HistoryPath, () => new List<RunHistoryEntry>());
    }

    public static void SaveHistory(IEnumerable<RunHistoryEntry> entries)
    {
        var paths = GetPaths();
        File.WriteAllText(paths.HistoryPath, JsonSerializer.Serialize(entries.ToList(), JsonOptions));
    }

    public static void AppendHistory(RunHistoryEntry entry)
    {
        var history = LoadHistory();
        history.Insert(0, entry);
        SaveHistory(history);
    }

    public static void UpdateHistory(Func<List<RunHistoryEntry>, List<RunHistoryEntry>> updater)
    {
        var updated = updater(LoadHistory());
        SaveHistory(updated);
    }

    public static void OpenPath(string path)
    {
        Process.Start(new ProcessStartInfo
        {
            FileName = path,
            UseShellExecute = true
        });
    }

    public static void OpenShellUrl(string url)
    {
        if (TryOpenChromiumAppWindow(url))
        {
            return;
        }

        OpenPath(url);
    }

    private static void EnsureDataDirectoryMigrated(string baseDirectory, string legacyDirectory)
    {
        if (Directory.Exists(baseDirectory) || !Directory.Exists(legacyDirectory))
        {
            return;
        }

        try
        {
            CopyDirectory(legacyDirectory, baseDirectory);
        }
        catch
        {
            Directory.CreateDirectory(baseDirectory);
        }
    }

    private static void CopyDirectory(string sourceDirectory, string destinationDirectory)
    {
        Directory.CreateDirectory(destinationDirectory);

        foreach (var file in Directory.EnumerateFiles(sourceDirectory))
        {
            var destinationFile = Path.Combine(destinationDirectory, Path.GetFileName(file));
            File.Copy(file, destinationFile, overwrite: false);
        }

        foreach (var directory in Directory.EnumerateDirectories(sourceDirectory))
        {
            var destinationSubdirectory = Path.Combine(destinationDirectory, Path.GetFileName(directory));
            CopyDirectory(directory, destinationSubdirectory);
        }
    }

    private static T LoadJsonWithFallback<T>(string path, Func<T> fallbackFactory, Func<T, T>? normalize = null)
    {
        if (!File.Exists(path))
        {
            return fallbackFactory();
        }

        try
        {
            var value = JsonSerializer.Deserialize<T>(File.ReadAllText(path), JsonOptions);
            if (value is null)
            {
                BackupInvalidJsonFile(path);
                return fallbackFactory();
            }

            return normalize is null ? value : normalize(value);
        }
        catch (JsonException)
        {
            BackupInvalidJsonFile(path);
            return fallbackFactory();
        }
    }

    private static void BackupInvalidJsonFile(string path)
    {
        try
        {
            if (!File.Exists(path))
            {
                return;
            }

            var directory = Path.GetDirectoryName(path);
            if (string.IsNullOrWhiteSpace(directory))
            {
                return;
            }

            var baseName = Path.GetFileNameWithoutExtension(path);
            var extension = Path.GetExtension(path);
            var backupPath = Path.Combine(directory, $"{baseName}.invalid-{DateTimeOffset.Now:yyyyMMdd-HHmmss}{extension}");
            File.Move(path, backupPath, overwrite: true);
        }
        catch
        {
            // Fall back to defaults even if the bad file cannot be renamed.
        }
    }

    private static RulesConfig CreateDefaultRules()
    {
        return new RulesConfig
        {
            CategoryRules = new List<CategoryRule>
            {
                new() { Name = "Images", Extensions = new List<string> { ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".webp", ".svg", ".heic", ".ico", ".raw", ".cr2", ".nef" } },
                new() { Name = "Videos", Extensions = new List<string> { ".mp4", ".mov", ".avi", ".mkv", ".wmv", ".webm", ".flv", ".mpeg", ".mpg", ".m4v" } },
                new() { Name = "Audio", Extensions = new List<string> { ".mp3", ".wav", ".flac", ".aac", ".ogg", ".m4a", ".wma", ".aiff", ".alac" } },
                new() { Name = "Documents", Extensions = new List<string> { ".pdf", ".doc", ".docx", ".txt", ".rtf", ".odt", ".md", ".tex" } },
                new() { Name = "Spreadsheets", Extensions = new List<string> { ".xls", ".xlsx", ".csv", ".tsv", ".ods" } },
                new() { Name = "Presentations", Extensions = new List<string> { ".ppt", ".pptx", ".odp", ".key" } },
                new() { Name = "Archives", Extensions = new List<string> { ".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz", ".iso" } },
                new() { Name = "Code", Extensions = new List<string> { ".ps1", ".psm1", ".psd1", ".bat", ".cmd", ".py", ".js", ".jsx", ".ts", ".tsx", ".html", ".css", ".scss", ".json", ".xml", ".yaml", ".yml", ".sql", ".java", ".c", ".cpp", ".cs", ".go", ".rs", ".php", ".rb", ".sh", ".ipynb" } },
                new() { Name = "Design", Extensions = new List<string> { ".psd", ".ai", ".fig", ".sketch", ".xd" } },
                new() { Name = "Fonts", Extensions = new List<string> { ".ttf", ".otf", ".woff", ".woff2" } },
                new() { Name = "Ebooks", Extensions = new List<string> { ".epub", ".mobi", ".azw3" } },
                new() { Name = "Executables", Extensions = new List<string> { ".exe", ".msi", ".msix", ".appx" } },
                new() { Name = "Shortcuts", Extensions = new List<string> { ".lnk", ".url" } }
            },
            KeywordRules = new List<KeywordRule>
            {
                new() { Name = "Invoices", Keywords = new List<string> { "invoice", "receipt", "bill" } },
                new() { Name = "Contracts", Keywords = new List<string> { "contract", "agreement", "nda" } },
                new() { Name = "Screenshots", Keywords = new List<string> { "screenshot", "screen shot", "capture" } }
            },
            ExcludeRules = new ExcludeRules
            {
                IgnoreHiddenFiles = true,
                IgnoreSystemFiles = true,
                Extensions = new List<string> { ".tmp", ".bak", ".part" },
                FolderNames = new List<string> { "_organized", ".git", "node_modules", "bin", "obj" }
            },
            NamingRules = new NamingRules
            {
                FileNameTemplate = "{originalName}",
                InvalidCharacterReplacement = "_"
            },
            SizeBands = new List<SizeBandRule>
            {
                new() { Name = "Tiny", MaxBytes = 1023 },
                new() { Name = "Small", MaxBytes = 10L * 1024 * 1024 - 1 },
                new() { Name = "Medium", MaxBytes = 100L * 1024 * 1024 - 1 },
                new() { Name = "Large", MaxBytes = 1024L * 1024 * 1024 - 1 },
                new() { Name = "Huge", MaxBytes = null }
            }
        };
    }

    private static void NormalizeRules(RulesConfig config)
    {
        foreach (var category in config.CategoryRules)
        {
            category.Name = (category.Name ?? string.Empty).Trim();
            category.Extensions = category.Extensions
                .Where(extension => !string.IsNullOrWhiteSpace(extension))
                .Select(NormalizeExtension)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        foreach (var keyword in config.KeywordRules)
        {
            keyword.Name = (keyword.Name ?? string.Empty).Trim();
            keyword.Keywords = keyword.Keywords
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        config.ExcludeRules.Extensions = config.ExcludeRules.Extensions
            .Where(extension => !string.IsNullOrWhiteSpace(extension))
            .Select(NormalizeExtension)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        config.ExcludeRules.FolderNames = config.ExcludeRules.FolderNames
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Select(name => name.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        config.NamingRules.FileNameTemplate = string.IsNullOrWhiteSpace(config.NamingRules.FileNameTemplate)
            ? "{originalName}"
            : config.NamingRules.FileNameTemplate;

        config.NamingRules.InvalidCharacterReplacement = string.IsNullOrEmpty(config.NamingRules.InvalidCharacterReplacement)
            ? "_"
            : config.NamingRules.InvalidCharacterReplacement;

        if (config.SizeBands.Count == 0)
        {
            config.SizeBands = CreateDefaultRules().SizeBands;
        }
    }

    private static bool TryOpenChromiumAppWindow(string url)
    {
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) ||
            !(uri.Scheme.Equals(Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) ||
              uri.Scheme.Equals(Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)))
        {
            return false;
        }

        foreach (var browserPath in GetChromiumCandidates())
        {
            if (!File.Exists(browserPath))
            {
                continue;
            }

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = browserPath,
                    Arguments = $"--app=\"{url}\" --new-window --window-size=1260,860",
                    UseShellExecute = false
                });

                return true;
            }
            catch
            {
                // Try the next installed browser candidate.
            }
        }

        return false;
    }

    private static IEnumerable<string> GetChromiumCandidates()
    {
        yield return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Microsoft", "Edge", "Application", "msedge.exe");
        yield return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Microsoft", "Edge", "Application", "msedge.exe");
        yield return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "BraveSoftware", "Brave-Browser", "Application", "brave.exe");
        yield return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "BraveSoftware", "Brave-Browser", "Application", "brave.exe");
        yield return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BraveSoftware", "Brave-Browser", "Application", "brave.exe");
        yield return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Google", "Chrome", "Application", "chrome.exe");
        yield return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Google", "Chrome", "Application", "chrome.exe");
    }

    private static string NormalizeExtension(string extension)
    {
        var trimmed = extension.Trim();
        return trimmed.StartsWith(".", StringComparison.Ordinal) ? trimmed.ToLowerInvariant() : "." + trimmed.ToLowerInvariant();
    }
}
