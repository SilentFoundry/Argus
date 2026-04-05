using System;
using System.Collections.Generic;

namespace FileOrganizer;

internal enum OperationMode
{
    Move,
    Copy
}

internal enum DuplicateHandlingMode
{
    Rename,
    Skip,
    Overwrite,
    KeepNewest,
    CompareByHash
}

internal enum OrganizationMode
{
    ByType,
    ByExtension,
    ByYearMonth,
    ByDateModified,
    BySize,
    ByKeyword,
    CategoryYearMonth
}

internal enum FileOperationStatus
{
    Planned,
    Executed,
    Skipped,
    Failed,
    Undone
}

internal enum RunKind
{
    Organize,
    Undo
}

internal enum ReportExportFormat
{
    Text,
    Json
}

internal enum PdfRenameMode
{
    Disabled,
    TitleOnly,
    ContentHeadingOnly,
    TitleThenContent
}

internal enum ConversionStatus
{
    Planned,
    Executed,
    Skipped,
    Failed
}

internal sealed record OrganizerOptions
{
    public string? SourcePath { get; init; }
    public string? DestinationPath { get; init; }
    public string? TagManifestPath { get; init; }
    public string? RulesPath { get; init; }
    public OperationMode Mode { get; init; } = OperationMode.Move;
    public DuplicateHandlingMode DuplicateHandling { get; init; } = DuplicateHandlingMode.Rename;
    public OrganizationMode OrganizationMode { get; init; } = OrganizationMode.ByType;
    public bool TopLevelOnly { get; init; }
    public bool UseExtensionSubfolders { get; init; }
    public bool WriteTagManifest { get; init; }
    public bool WhatIf { get; init; }
    public bool CleanEmptyFolders { get; init; }
    public PdfRenameMode PdfRenameMode { get; init; } = PdfRenameMode.TitleThenContent;
    public IReadOnlyList<string> IncludeExtensions { get; init; } = Array.Empty<string>();
    public bool IncludeConvert { get; init; }
    public string ConvertSourceFormat { get; init; } = "Any";
    public string ConvertTargetFormat { get; init; } = "Keep";
    public string ConvertQuality { get; init; } = "Balanced";
    public bool MediaMusicUseMetadata { get; init; } = true;
    public bool MediaImagesUseMetadata { get; init; } = true;
    public bool MediaVideoUseMetadata { get; init; } = true;
}

internal sealed record OrganizerProgressInfo(
    string Phase,
    int Completed,
    int Total,
    string Message,
    bool IsIndeterminate = false);

internal sealed record FileActionRecord
{
    public required string SourcePath { get; init; }
    public required string DestinationPath { get; init; }
    public required string DestinationFolder { get; init; }
    public required string GroupName { get; init; }
    public required string Category { get; init; }
    public required string Extension { get; init; }
    public required long SizeBytes { get; init; }
    public required string LastModified { get; init; }
    public required string[] Tags { get; init; }
    public required DuplicateHandlingMode DuplicateHandling { get; init; }
    public bool DestinationOriginallyExisted { get; init; }
    public string? ExistingDestinationPath { get; init; }
    public bool ShouldExecute { get; init; } = true;
    public string? SkipReason { get; init; }
    public string? BackupPath { get; init; }
    public string? KeywordGroup { get; init; }
    public FileOperationStatus Status { get; init; } = FileOperationStatus.Planned;
    public string? StatusMessage { get; init; }
    public IReadOnlyList<string> GeneratedPaths { get; init; } = Array.Empty<string>();
}

internal sealed record ConversionRecord
{
    public required string SourcePath { get; init; }
    public required string OutputPath { get; init; }
    public required string TargetFormat { get; init; }
    public required string Converter { get; init; }
    public ConversionStatus Status { get; init; } = ConversionStatus.Planned;
    public string? Message { get; init; }
}

internal sealed record OrganizerRunResult
{
    public required string RunId { get; init; }
    public required DateTimeOffset Timestamp { get; init; }
    public required RunKind RunKind { get; init; }
    public required string SourceRoot { get; init; }
    public required string DestinationRoot { get; init; }
    public required string RulesPath { get; init; }
    public required string TagManifestPath { get; init; }
    public required string LogFilePath { get; init; }
    public required string OperationLogPath { get; init; }
    public required OperationMode Mode { get; init; }
    public required DuplicateHandlingMode DuplicateHandling { get; init; }
    public required OrganizationMode OrganizationMode { get; init; }
    public PdfRenameMode PdfRenameMode { get; init; } = PdfRenameMode.TitleThenContent;
    public bool WhatIf { get; init; }
    public bool WriteTagManifest { get; init; }
    public bool CleanEmptyFolders { get; init; }
    public bool IncludeConvert { get; init; }
    public string ConvertSourceFormat { get; init; } = "Any";
    public string ConvertTargetFormat { get; init; } = "Keep";
    public string ConvertQuality { get; init; } = "Balanced";
    public required IReadOnlyList<FileActionRecord> Operations { get; init; }
    public required IReadOnlyList<ConversionRecord> Conversions { get; init; }
    public required IReadOnlyList<string> CleanedDirectories { get; init; }
    public required string SummaryText { get; init; }
    public required string DetailedReport { get; init; }
}

internal sealed record UndoResult
{
    public required string RunId { get; init; }
    public required string TargetRunId { get; init; }
    public required string LogFilePath { get; init; }
    public required string SummaryText { get; init; }
    public required string DetailedReport { get; init; }
}

internal sealed record OperationLogDocument
{
    public required string RunId { get; init; }
    public required RunKind RunKind { get; init; }
    public required DateTimeOffset Timestamp { get; init; }
    public required string SourceRoot { get; init; }
    public required string DestinationRoot { get; init; }
    public required OperationMode Mode { get; init; }
    public required DuplicateHandlingMode DuplicateHandling { get; init; }
    public required OrganizationMode OrganizationMode { get; init; }
    public PdfRenameMode PdfRenameMode { get; init; } = PdfRenameMode.TitleThenContent;
    public required bool WhatIf { get; init; }
    public string? TagManifestPath { get; init; }
    public bool IncludeConvert { get; init; }
    public string ConvertSourceFormat { get; init; } = "Any";
    public string ConvertTargetFormat { get; init; } = "Keep";
    public string ConvertQuality { get; init; } = "Balanced";
    public required IReadOnlyList<FileActionRecord> Operations { get; init; }
    public required IReadOnlyList<ConversionRecord> Conversions { get; init; }
    public required IReadOnlyList<string> CleanedDirectories { get; init; }
}

internal sealed record RunHistoryEntry
{
    public required string RunId { get; init; }
    public required RunKind RunKind { get; init; }
    public required DateTimeOffset Timestamp { get; init; }
    public required string SourceRoot { get; init; }
    public required string DestinationRoot { get; init; }
    public required OperationMode Mode { get; init; }
    public required DuplicateHandlingMode DuplicateHandling { get; init; }
    public required OrganizationMode OrganizationMode { get; init; }
    public PdfRenameMode PdfRenameMode { get; init; } = PdfRenameMode.TitleThenContent;
    public required bool WhatIf { get; init; }
    public required int PlannedCount { get; init; }
    public required int ExecutedCount { get; init; }
    public required int SkippedCount { get; init; }
    public int ConversionCount { get; init; }
    public required string RulesPath { get; init; }
    public string? TagManifestPath { get; init; }
    public required string LogFilePath { get; init; }
    public required string OperationLogPath { get; init; }
    public string? SummaryText { get; init; }
    public bool CanUndo { get; init; }
    public bool IsUndone { get; init; }
    public string? UndoneByRunId { get; init; }
}

internal sealed record AppSettings
{
    public string? LastSourcePath { get; init; }
    public string? LastDestinationPath { get; init; }
    public string? LastTagManifestPath { get; init; }
    public string? LastRulesPath { get; init; }
    public OperationMode Mode { get; init; } = OperationMode.Move;
    public DuplicateHandlingMode DuplicateHandling { get; init; } = DuplicateHandlingMode.Rename;
    public OrganizationMode OrganizationMode { get; init; } = OrganizationMode.ByType;
    public bool TopLevelOnly { get; init; }
    public bool UseExtensionSubfolders { get; init; }
    public bool WriteTagManifest { get; init; }
    public bool CleanEmptyFolders { get; init; }
    public bool OpenDestinationWhenDone { get; init; } = true;
    public bool ExportReport { get; init; }
    public ReportExportFormat ReportExportFormat { get; init; } = ReportExportFormat.Text;
    public bool CopySummaryToClipboard { get; init; }
    public PdfRenameMode PdfRenameMode { get; init; } = PdfRenameMode.TitleThenContent;
    public bool IncludeConvert { get; init; }
    public string ConvertSourceFormat { get; init; } = "Any";
    public string ConvertTargetFormat { get; init; } = "Keep";
    public string ConvertQuality { get; init; } = "Balanced";
    public bool MediaMusicUseMetadata { get; init; } = true;
    public bool MediaImagesUseMetadata { get; init; } = true;
    public bool MediaVideoUseMetadata { get; init; } = true;
    public int WindowWidth { get; init; } = 1280;
    public int WindowHeight { get; init; } = 860;
    public int WindowLayoutVersion { get; init; }
    public int SelectedTabIndex { get; init; }
}

internal sealed record AppPaths(
    string BaseDirectory,
    string SettingsPath,
    string HistoryPath,
    string LogsDirectory,
    string ReportsDirectory,
    string DefaultRulesPath);

internal sealed class RulesConfig
{
    public List<CategoryRule> CategoryRules { get; set; } = new();
    public List<KeywordRule> KeywordRules { get; set; } = new();
    public ExcludeRules ExcludeRules { get; set; } = new();
    public NamingRules NamingRules { get; set; } = new();
    public List<SizeBandRule> SizeBands { get; set; } = new();
}

internal sealed class CategoryRule
{
    public string Name { get; set; } = string.Empty;
    public List<string> Extensions { get; set; } = new();
}

internal sealed class KeywordRule
{
    public string Name { get; set; } = string.Empty;
    public List<string> Keywords { get; set; } = new();
}

internal sealed class ExcludeRules
{
    public bool IgnoreHiddenFiles { get; set; } = true;
    public bool IgnoreSystemFiles { get; set; } = true;
    public List<string> Extensions { get; set; } = new();
    public List<string> FolderNames { get; set; } = new();
    public long? MinFileSizeBytes { get; set; }
    public long? MaxFileSizeBytes { get; set; }
}

internal sealed class NamingRules
{
    public string FileNameTemplate { get; set; } = "{originalName}";
    public string InvalidCharacterReplacement { get; set; } = "_";
}

internal sealed class SizeBandRule
{
    public string Name { get; set; } = string.Empty;
    public long? MaxBytes { get; set; }
}

internal sealed record LoadedRules(string Path, RulesConfig Config);
