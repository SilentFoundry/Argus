using System;
using System.Collections.Generic;

namespace FileOrganizer;

internal sealed record UiStatePayload
{
    public string? SourcePath { get; init; }
    public string? DestinationPath { get; init; }
    public string? TagManifestPath { get; init; }
    public string? RulesPath { get; init; }
    public string? IncludeExtensions { get; init; }
    public OperationMode Mode { get; init; } = OperationMode.Move;
    public DuplicateHandlingMode DuplicateHandling { get; init; } = DuplicateHandlingMode.Rename;
    public OrganizationMode OrganizationMode { get; init; } = OrganizationMode.ByType;
    public PdfRenameMode PdfRenameMode { get; init; } = PdfRenameMode.TitleThenContent;
    public bool TopLevelOnly { get; init; }
    public bool UseExtensionSubfolders { get; init; }
    public bool WriteTagManifest { get; init; }
    public bool CleanEmptyFolders { get; init; }
    public bool OpenDestinationWhenDone { get; init; } = true;
    public bool ExportReport { get; init; }
    public ReportExportFormat ReportExportFormat { get; init; } = ReportExportFormat.Text;
    public bool CopySummaryToClipboard { get; init; }
    public bool IncludeConvert { get; init; }
    public string ConvertSourceFormat { get; init; } = "Any";
    public string ConvertTargetFormat { get; init; } = "Keep";
    public string ConvertQuality { get; init; } = "Balanced";
    public bool MediaMusicUseMetadata { get; init; } = true;
    public bool MediaImagesUseMetadata { get; init; } = true;
    public bool MediaVideoUseMetadata { get; init; } = true;
    public int SelectedTabIndex { get; init; }
}

internal sealed record UiUndoRequest
{
    public string? RunId { get; init; }
}

internal sealed record BrowserDialogRequest
{
    public string? InitialPath { get; init; }
}

internal sealed record OpenPathRequest
{
    public string? Path { get; init; }
}

internal sealed record PathResponse(string? Path);

internal sealed record FileContentResponse(string Path, string Contents);

internal sealed record LogFileItem(string Path, string FileName, DateTimeOffset LastWriteTime);

internal sealed record UiPathsResponse(
    string AppDirectory,
    string DataDirectory,
    string SettingsPath,
    string HistoryPath,
    string LogsDirectory,
    string ReportsDirectory,
    string DefaultRulesPath);

internal sealed record UiMetaResponse(
    string ProductName,
    string Version,
    string HostUrl,
    string HostMode);

internal sealed record UiCapabilitiesResponse(
    bool ImagePipelineAvailable,
    bool FfmpegAvailable,
    bool LibreOfficeAvailable,
    bool ImageMagickAvailable,
    string AudioVideoPipeline,
    string DocumentPipeline,
    string ImagePipeline,
    string ShellMode);

internal sealed record UiStateResponse(
    UiMetaResponse Meta,
    UiPathsResponse Paths,
    UiCapabilitiesResponse Capabilities,
    AppSettings Settings,
    IReadOnlyList<RunHistoryEntry> History,
    IReadOnlyList<LogFileItem> Logs,
    string RulesPreviewPath,
    string RulesPreview);

internal sealed record UiOperationResponse(
    OrganizerRunResult Result,
    UiStateResponse State,
    string? ExportedReportPath);

internal sealed record UiUndoResponse(
    UndoResult Result,
    UiStateResponse State);

internal sealed record ApiErrorResponse(string Message);
