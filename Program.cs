using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

namespace FileOrganizer;

internal static class Program
{
    [STAThread]
    private static async Task<int> Main(string[] args)
    {
        var launchOptions = WebLaunchOptions.Parse(args);
        if (!launchOptions.StartWeb)
        {
            return CliMode.Run(args);
        }

                var hostUrl = $"http://127.0.0.1:{launchOptions.Port ?? GetAvailablePort()}";

        var builder = WebApplication.CreateBuilder(new WebApplicationOptions
        {
            Args = Array.Empty<string>(),
            ContentRootPath = AppContext.BaseDirectory,
            WebRootPath = Path.Combine(AppContext.BaseDirectory, "wwwroot")
        });
        builder.WebHost.UseUrls(hostUrl);
        builder.Services.ConfigureHttpJsonOptions(options =>
        {
            options.SerializerOptions.Converters.Add(new JsonStringEnumConverter());
            options.SerializerOptions.WriteIndented = false;
        });
        builder.Services.AddSingleton<HeartbeatService>();

        var app = builder.Build();
        app.UseDefaultFiles();
        app.UseStaticFiles();
        MapApi(app, hostUrl);
        app.MapFallbackToFile("index.html");

        await app.StartAsync();

        if (!launchOptions.NoBrowser)
        {
            AppStorage.OpenShellUrl(hostUrl);
        }

        await app.WaitForShutdownAsync();
        return 0;
    }

    private static void MapApi(WebApplication app, string hostUrl)
    {
        app.MapGet("/api/state", (HeartbeatService heartbeat) =>
        {
            heartbeat.Beat();
            return Results.Ok(BuildStateResponse(hostUrl));
        });

        app.MapGet("/api/file-content", (string path, HeartbeatService heartbeat) =>
        {
            heartbeat.Beat();
            return TryReadFile(path);
        });

        app.MapGet("/api/rules-content", (string? path, HeartbeatService heartbeat) =>
        {
            heartbeat.Beat();
            var resolvedPath = string.IsNullOrWhiteSpace(path)
                ? AppStorage.LoadSettings().LastRulesPath ?? AppStorage.GetPaths().DefaultRulesPath
                : path;

            return TryReadFile(resolvedPath);
        });

        app.MapPost("/api/settings", (UiStatePayload payload, HeartbeatService heartbeat) =>
        {
            heartbeat.Beat();
            AppStorage.SaveSettings(BuildSavedSettings(payload));
            return Results.Ok(BuildStateResponse(hostUrl));
        });

        app.MapPost("/api/dialog/source", (BrowserDialogRequest request, HeartbeatService heartbeat) =>
        {
            heartbeat.Beat();
            return Results.Ok(new PathResponse(DialogService.PickFolder(request.InitialPath, "Choose the folder to organize")));
        });

        app.MapPost("/api/dialog/destination", (BrowserDialogRequest request, HeartbeatService heartbeat) =>
        {
            heartbeat.Beat();
            return Results.Ok(new PathResponse(DialogService.PickFolder(request.InitialPath, "Choose the destination folder")));
        });

        app.MapPost("/api/dialog/manifest", (BrowserDialogRequest request, HeartbeatService heartbeat) =>
        {
            heartbeat.Beat();
            return Results.Ok(new PathResponse(DialogService.PickSaveFile(
                request.InitialPath,
                "Choose where to write the tag manifest",
                "JSON files (*.json)|*.json|All files (*.*)|*.*",
                "file-tags.json")));
        });

        app.MapPost("/api/dialog/rules", (BrowserDialogRequest request, HeartbeatService heartbeat) =>
        {
            heartbeat.Beat();
            return Results.Ok(new PathResponse(DialogService.PickOpenFile(
                request.InitialPath,
                "Choose a rules file",
                "JSON files (*.json)|*.json|All files (*.*)|*.*")));
        });

        app.MapPost("/api/open-path", (OpenPathRequest request, HeartbeatService heartbeat) =>
        {
            heartbeat.Beat();
            if (string.IsNullOrWhiteSpace(request.Path))
            {
                return Results.BadRequest(new ApiErrorResponse("Path is required."));
            }

            try
            {
                var resolved = Path.GetFullPath(request.Path);
                if (!File.Exists(resolved) && !Directory.Exists(resolved))
                {
                    return Results.BadRequest(new ApiErrorResponse($"Path does not exist: {resolved}"));
                }

                AppStorage.OpenPath(resolved);
                return Results.Ok(new PathResponse(resolved));
            }
            catch (Exception ex)
            {
                return Results.BadRequest(new ApiErrorResponse(ex.Message));
            }
        });

        app.MapPost("/api/operations/preview", async (UiStatePayload payload, HeartbeatService heartbeat, CancellationToken cancellationToken) =>
        {
            heartbeat.Beat();
            return await ExecuteOrganizerAsync(payload, whatIf: true, hostUrl, cancellationToken);
        });

        app.MapPost("/api/operations/run", async (UiStatePayload payload, HeartbeatService heartbeat, CancellationToken cancellationToken) =>
        {
            heartbeat.Beat();
            return await ExecuteOrganizerAsync(payload, whatIf: false, hostUrl, cancellationToken);
        });

        app.MapPost("/api/operations/undo", async (UiUndoRequest request, HeartbeatService heartbeat, CancellationToken cancellationToken) =>
        {
            heartbeat.Beat();
            return await ExecuteUndoAsync(request, hostUrl, cancellationToken);
        });

        app.MapPost("/api/heartbeat", (HeartbeatService heartbeat) =>
        {
            heartbeat.Beat();
            return Results.Ok();
        });

        app.MapPost("/api/exit", (IHostApplicationLifetime lifetime) =>
        {
            lifetime.StopApplication();
            return Results.Ok();
        });
    }

    private static UiStateResponse BuildStateResponse(string hostUrl)
    {
        var appPaths = AppStorage.GetPaths();
        var settings = AppStorage.LoadSettings();
        var rulesPath = string.IsNullOrWhiteSpace(settings.LastRulesPath) ? appPaths.DefaultRulesPath : settings.LastRulesPath;
        var capabilities = ConversionService.DetectCapabilities();

        return new UiStateResponse(
            new UiMetaResponse(
                ProductName: "ARGUS // File Operations",
                Version: GetVersionString(),
                HostUrl: hostUrl,
                HostMode: "Browser shell"),
            new UiPathsResponse(
                AppDirectory: AppContext.BaseDirectory,
                DataDirectory: appPaths.BaseDirectory,
                SettingsPath: appPaths.SettingsPath,
                HistoryPath: appPaths.HistoryPath,
                LogsDirectory: appPaths.LogsDirectory,
                ReportsDirectory: appPaths.ReportsDirectory,
                DefaultRulesPath: appPaths.DefaultRulesPath),
            new UiCapabilitiesResponse(
                capabilities.ImagePipelineAvailable,
                capabilities.FfmpegAvailable,
                capabilities.LibreOfficeAvailable,
                capabilities.ImageMagickAvailable,
                capabilities.AudioVideoPipeline,
                capabilities.DocumentPipeline,
                capabilities.ImagePipeline,
                capabilities.ShellMode),
            settings,
            LoadHistory(),
            LoadLogs(appPaths.LogsDirectory),
            rulesPath,
            TryReadFileContents(rulesPath));
    }

    private static IReadOnlyList<RunHistoryEntry> LoadHistory()
    {
        return AppStorage.LoadHistory()
            .OrderByDescending(entry => entry.Timestamp)
            .ToList();
    }

    private static IReadOnlyList<LogFileItem> LoadLogs(string logsDirectory)
    {
        if (!Directory.Exists(logsDirectory))
        {
            return Array.Empty<LogFileItem>();
        }

        return Directory.EnumerateFiles(logsDirectory, "*.log", SearchOption.AllDirectories)
            .Select(path => new FileInfo(path))
            .OrderByDescending(file => file.LastWriteTimeUtc)
            .Select(file => new LogFileItem(
                Path: file.FullName,
                FileName: file.Name,
                LastWriteTime: new DateTimeOffset(file.LastWriteTimeUtc)))
            .ToList();
    }

    private static async Task<IResult> ExecuteOrganizerAsync(UiStatePayload payload, bool whatIf, string hostUrl, CancellationToken cancellationToken)
    {
        try
        {
            AppStorage.SaveSettings(BuildSavedSettings(payload));

            var result = await Task.Run(
                () => OrganizerCore.Execute(BuildOrganizerOptions(payload, whatIf), progress: null, cancellationToken),
                cancellationToken);

            var exportedReportPath = payload.ExportReport
                ? await ExportReportAsync(result, payload.ReportExportFormat, cancellationToken)
                : null;

            if (!whatIf && payload.OpenDestinationWhenDone && Directory.Exists(result.DestinationRoot))
            {
                AppStorage.OpenPath(result.DestinationRoot);
            }

            return Results.Ok(new UiOperationResponse(result, BuildStateResponse(hostUrl), exportedReportPath));
        }
        catch (OperationCanceledException)
        {
            return Results.BadRequest(new ApiErrorResponse("Operation canceled."));
        }
        catch (Exception ex)
        {
            return Results.BadRequest(new ApiErrorResponse(ex.Message));
        }
    }

    private static async Task<IResult> ExecuteUndoAsync(UiUndoRequest request, string hostUrl, CancellationToken cancellationToken)
    {
        try
        {
            var result = await Task.Run(
                () => string.IsNullOrWhiteSpace(request.RunId)
                    ? UndoService.UndoLastRun(progress: null, cancellationToken)
                    : UndoService.UndoRun(request.RunId, progress: null, cancellationToken),
                cancellationToken);

            return Results.Ok(new UiUndoResponse(result, BuildStateResponse(hostUrl)));
        }
        catch (OperationCanceledException)
        {
            return Results.BadRequest(new ApiErrorResponse("Undo canceled."));
        }
        catch (Exception ex)
        {
            return Results.BadRequest(new ApiErrorResponse(ex.Message));
        }
    }

    private static OrganizerOptions BuildOrganizerOptions(UiStatePayload payload, bool whatIf)
    {
        return new OrganizerOptions
        {
            SourcePath = NormalizeInput(payload.SourcePath),
            DestinationPath = NormalizeInput(payload.DestinationPath),
            TagManifestPath = NormalizeInput(payload.TagManifestPath),
            RulesPath = NormalizeInput(payload.RulesPath),
            IncludeExtensions = ParseExtensions(payload.IncludeExtensions),
            Mode = payload.Mode,
            DuplicateHandling = payload.DuplicateHandling,
            OrganizationMode = payload.OrganizationMode,
            PdfRenameMode = payload.PdfRenameMode,
            TopLevelOnly = payload.TopLevelOnly,
            UseExtensionSubfolders = payload.UseExtensionSubfolders,
            WriteTagManifest = payload.WriteTagManifest,
            WhatIf = whatIf,
            CleanEmptyFolders = payload.CleanEmptyFolders,
            IncludeConvert = payload.IncludeConvert,
            ConvertSourceFormat = payload.ConvertSourceFormat,
            ConvertTargetFormat = payload.ConvertTargetFormat,
            ConvertQuality = payload.ConvertQuality,
            MediaMusicUseMetadata = payload.MediaMusicUseMetadata,
            MediaImagesUseMetadata = payload.MediaImagesUseMetadata,
            MediaVideoUseMetadata = payload.MediaVideoUseMetadata
        };
    }

    private static AppSettings BuildSavedSettings(UiStatePayload payload)
    {
        var current = AppStorage.LoadSettings();
        return current with
        {
            LastSourcePath = NormalizeInput(payload.SourcePath),
            LastDestinationPath = NormalizeInput(payload.DestinationPath),
            LastTagManifestPath = NormalizeInput(payload.TagManifestPath),
            LastRulesPath = NormalizeInput(payload.RulesPath) ?? AppStorage.GetPaths().DefaultRulesPath,
            Mode = payload.Mode,
            DuplicateHandling = payload.DuplicateHandling,
            OrganizationMode = payload.OrganizationMode,
            PdfRenameMode = payload.PdfRenameMode,
            TopLevelOnly = payload.TopLevelOnly,
            UseExtensionSubfolders = payload.UseExtensionSubfolders,
            WriteTagManifest = payload.WriteTagManifest,
            CleanEmptyFolders = payload.CleanEmptyFolders,
            OpenDestinationWhenDone = payload.OpenDestinationWhenDone,
            ExportReport = payload.ExportReport,
            ReportExportFormat = payload.ReportExportFormat,
            CopySummaryToClipboard = payload.CopySummaryToClipboard,
            IncludeConvert = payload.IncludeConvert,
            ConvertSourceFormat = payload.ConvertSourceFormat,
            ConvertTargetFormat = payload.ConvertTargetFormat,
            ConvertQuality = payload.ConvertQuality,
            MediaMusicUseMetadata = payload.MediaMusicUseMetadata,
            MediaImagesUseMetadata = payload.MediaImagesUseMetadata,
            MediaVideoUseMetadata = payload.MediaVideoUseMetadata,
            SelectedTabIndex = payload.SelectedTabIndex
        };
    }

    private static async Task<string> ExportReportAsync(OrganizerRunResult result, ReportExportFormat format, CancellationToken cancellationToken)
    {
        var appPaths = AppStorage.GetPaths();
        Directory.CreateDirectory(appPaths.ReportsDirectory);

        var reportPath = Path.Combine(
            appPaths.ReportsDirectory,
            format == ReportExportFormat.Json
                ? $"{result.RunId}.json"
                : $"{result.RunId}.txt");

        if (format == ReportExportFormat.Json)
        {
            var payload = new
            {
                result.RunId,
                result.Timestamp,
                result.SourceRoot,
                result.DestinationRoot,
                result.Mode,
                result.DuplicateHandling,
                result.OrganizationMode,
                result.PdfRenameMode,
                result.WhatIf,
                result.WriteTagManifest,
                result.CleanEmptyFolders,
                result.IncludeConvert,
                result.ConvertSourceFormat,
                result.ConvertTargetFormat,
                result.ConvertQuality,
                result.TagManifestPath,
                result.LogFilePath,
                result.OperationLogPath,
                result.Operations,
                result.Conversions,
                result.CleanedDirectories,
                result.SummaryText
            };

            await File.WriteAllTextAsync(reportPath, JsonSerializer.Serialize(payload, new JsonSerializerOptions { WriteIndented = true }), cancellationToken);
        }
        else
        {
            await File.WriteAllTextAsync(reportPath, result.DetailedReport, cancellationToken);
        }

        return reportPath;
    }

    private static IResult TryReadFile(string path)
    {
        try
        {
            var resolved = Path.GetFullPath(path);
            if (!File.Exists(resolved))
            {
                return Results.BadRequest(new ApiErrorResponse($"File not found: {resolved}"));
            }

            return Results.Ok(new FileContentResponse(resolved, File.ReadAllText(resolved)));
        }
        catch (Exception ex)
        {
            return Results.BadRequest(new ApiErrorResponse(ex.Message));
        }
    }

    private static string TryReadFileContents(string path)
    {
        try
        {
            var resolved = Path.GetFullPath(path);
            return File.Exists(resolved) ? File.ReadAllText(resolved) : string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string? NormalizeInput(string? value)
    {
        var trimmed = value?.Trim();
        return string.IsNullOrWhiteSpace(trimmed) ? null : trimmed;
    }

    private static IReadOnlyList<string> ParseExtensions(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return Array.Empty<string>();
        }

        return value
            .Split(new[] { ',', ';', '\r', '\n', '\t', ' ' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(part => part.Trim())
            .Where(part => !string.IsNullOrWhiteSpace(part))
            .Select(part => part.StartsWith('.') ? part.ToLowerInvariant() : $".{part.ToLowerInvariant()}")
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private static string GetVersionString()
    {
        var version = Assembly.GetEntryAssembly()?.GetName().Version;
        return version is null ? "0.0.0" : $"{version.Major}.{version.Minor}.{version.Build}";
    }

    private static int GetAvailablePort()
    {
        using var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        var port = ((IPEndPoint)listener.LocalEndpoint).Port;
        listener.Stop();
        return port;
    }

    private sealed record WebLaunchOptions(bool StartWeb, bool NoBrowser, int? Port)
    {
        public static WebLaunchOptions Parse(string[] args)
        {
            var startWeb = args.Length == 0 ||
                           args.Any(argument => string.Equals(argument, "--web", StringComparison.OrdinalIgnoreCase) ||
                                                string.Equals(argument, "--serve", StringComparison.OrdinalIgnoreCase));

            if (!startWeb)
            {
                return new WebLaunchOptions(false, false, null);
            }

            var noBrowser = args.Any(argument => string.Equals(argument, "--no-browser", StringComparison.OrdinalIgnoreCase));
            int? port = null;

            for (var index = 0; index < args.Length; index++)
            {
                if (!string.Equals(args[index], "--port", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (index + 1 >= args.Length || !int.TryParse(args[index + 1], out var parsedPort))
                {
                    throw new InvalidOperationException("Missing or invalid value for --port.");
                }

                port = parsedPort;
                index++;
            }

            return new WebLaunchOptions(true, noBrowser, port);
        }
    }
}







