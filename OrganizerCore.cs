using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading;

namespace FileOrganizer;

internal static class OrganizerCore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    public static OrganizerRunResult Execute(
        OrganizerOptions options,
        IProgress<OrganizerProgressInfo>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var appPaths = AppStorage.GetPaths();
        var loadedRules = AppStorage.LoadRules(options.RulesPath);

        if (string.IsNullOrWhiteSpace(options.SourcePath))
        {
            throw new InvalidOperationException("SourcePath is required.");
        }

        var sourceRoot = Path.GetFullPath(options.SourcePath);
        if (!Directory.Exists(sourceRoot))
        {
            throw new DirectoryNotFoundException($"SourcePath does not exist or is not a folder: {options.SourcePath}");
        }

        var destinationRoot = string.IsNullOrWhiteSpace(options.DestinationPath)
            ? Path.Combine(sourceRoot, "_organized")
            : Path.GetFullPath(options.DestinationPath);

        if (PathsEqual(sourceRoot, destinationRoot))
        {
            throw new InvalidOperationException("DestinationPath must be different from SourcePath.");
        }

        var tagManifestPath = string.IsNullOrWhiteSpace(options.TagManifestPath)
            ? Path.Combine(destinationRoot, "file-tags.json")
            : Path.GetFullPath(options.TagManifestPath);

        var runId = $"run-{DateTimeOffset.Now:yyyyMMdd-HHmmss}-{Guid.NewGuid():N}";
        var runDirectory = Path.Combine(appPaths.LogsDirectory, runId);
        Directory.CreateDirectory(runDirectory);

        if (!options.WhatIf)
        {
            Directory.CreateDirectory(destinationRoot);
        }

        var candidateFiles = GetCandidateFiles(sourceRoot, destinationRoot, loadedRules.Config, options.TopLevelOnly, options.IncludeExtensions, progress, cancellationToken);
        var operations = BuildPlan(candidateFiles, destinationRoot, loadedRules.Config, options, progress, cancellationToken);
        var executedOperations = ExecutePlan(operations, options, runDirectory, progress, cancellationToken);
        var conversionBatch = ConversionService.Process(executedOperations, options, progress, cancellationToken);
        executedOperations = conversionBatch.Operations.ToList();

        if (options.WriteTagManifest && !options.WhatIf && executedOperations.Any(item => item.Status == FileOperationStatus.Executed))
        {
            WriteTagManifest(executedOperations.Where(item => item.Status == FileOperationStatus.Executed), tagManifestPath);
        }

        var cleanedDirectories = (!options.WhatIf && options.Mode == OperationMode.Move && options.CleanEmptyFolders)
            ? CleanupEmptyDirectories(sourceRoot, destinationRoot, progress, cancellationToken)
            : new List<string>();

        var operationLogPath = Path.Combine(runDirectory, $"{runId}.operations.json");
        var logFilePath = Path.Combine(runDirectory, $"{runId}.log");

        var result = new OrganizerRunResult
        {
            RunId = runId,
            Timestamp = DateTimeOffset.Now,
            RunKind = RunKind.Organize,
            SourceRoot = sourceRoot,
            DestinationRoot = destinationRoot,
            RulesPath = loadedRules.Path,
            TagManifestPath = tagManifestPath,
            LogFilePath = logFilePath,
            OperationLogPath = operationLogPath,
            Mode = options.Mode,
            DuplicateHandling = options.DuplicateHandling,
            OrganizationMode = options.OrganizationMode,
            PdfRenameMode = options.PdfRenameMode,
            WhatIf = options.WhatIf,
            WriteTagManifest = options.WriteTagManifest,
            CleanEmptyFolders = options.CleanEmptyFolders,
            IncludeConvert = options.IncludeConvert,
            ConvertSourceFormat = options.ConvertSourceFormat,
            ConvertTargetFormat = options.ConvertTargetFormat,
            ConvertQuality = options.ConvertQuality,
            Operations = executedOperations,
            Conversions = conversionBatch.Conversions,
            CleanedDirectories = cleanedDirectories,
            SummaryText = string.Empty,
            DetailedReport = string.Empty
        };

        result = result with
        {
            SummaryText = BuildSummary(result),
            DetailedReport = BuildDetailedReport(result)
        };

        WriteLogArtifacts(result);
        AppendHistoryEntry(result);

        progress?.Report(new OrganizerProgressInfo("Completed", executedOperations.Count(item => item.Status == FileOperationStatus.Executed), executedOperations.Count, "Completed"));
        return result;
    }

    public static string BuildSummary(OrganizerRunResult result)
    {
        var plannedCount = result.Operations.Count;
        var executedCount = result.Operations.Count(item => item.Status == FileOperationStatus.Executed);
        var skippedCount = result.Operations.Count(item => item.Status == FileOperationStatus.Skipped);
        var failedCount = result.Operations.Count(item => item.Status == FileOperationStatus.Failed);

        var groupCounts = result.Operations
            .Where(item => item.ShouldExecute || item.Status == FileOperationStatus.Executed)
            .GroupBy(item => item.GroupName, StringComparer.OrdinalIgnoreCase)
            .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var builder = new StringBuilder();
        builder.AppendLine($"Run ID: {result.RunId}");
        builder.AppendLine($"Planned files: {plannedCount}");
        builder.AppendLine($"Changed files: {executedCount}");
        builder.AppendLine($"Skipped files: {skippedCount}");
        builder.AppendLine($"Failed files: {failedCount}");
        builder.AppendLine($"Mode: {result.Mode}");
        builder.AppendLine($"Organization mode: {result.OrganizationMode}");
        builder.AppendLine($"Duplicate handling: {result.DuplicateHandling}");
        builder.AppendLine($"Metadata rename mode: {result.PdfRenameMode}");
        if (result.IncludeConvert)
        {
            builder.AppendLine($"Convert source filter: {result.ConvertSourceFormat}");
            builder.AppendLine($"Convert target: {result.ConvertTargetFormat}");
            builder.AppendLine($"Convert quality: {result.ConvertQuality}");
        }
        builder.AppendLine($"Source: {result.SourceRoot}");
        builder.AppendLine($"Destination: {result.DestinationRoot}");
        builder.AppendLine($"Rules: {result.RulesPath}");
        builder.AppendLine($"Log file: {result.LogFilePath}");
        builder.AppendLine($"History file: {AppStorage.GetPaths().HistoryPath}");

        if (result.WhatIf)
        {
            builder.AppendLine("Preview only: no files were changed.");
        }

        if (result.WriteTagManifest)
        {
            builder.AppendLine(result.WhatIf
                ? $"Tag manifest: would write {result.TagManifestPath}"
                : $"Tag manifest: {result.TagManifestPath}");
        }

        if (result.CleanedDirectories.Count > 0)
        {
            builder.AppendLine($"Cleaned empty folders: {result.CleanedDirectories.Count}");
        }

        if (result.Conversions.Count > 0)
        {
            builder.AppendLine($"Conversions: {result.Conversions.Count(item => item.Status == ConversionStatus.Executed)} executed / {result.Conversions.Count(item => item.Status == ConversionStatus.Planned)} planned / {result.Conversions.Count(item => item.Status == ConversionStatus.Skipped)} skipped / {result.Conversions.Count(item => item.Status == ConversionStatus.Failed)} failed");
        }

        if (plannedCount == 0)
        {
            builder.AppendLine("No files found to process.");
            return builder.ToString().TrimEnd();
        }

        builder.AppendLine();
        builder.AppendLine("Group summary:");
        foreach (var group in groupCounts)
        {
            builder.AppendLine($" - {group.Key}: {group.Count()}");
        }

        return builder.ToString().TrimEnd();
    }

    public static string BuildDetailedReport(OrganizerRunResult result)
    {
        var builder = new StringBuilder();
        builder.AppendLine(BuildSummary(result));

        if (result.Operations.Count > 0)
        {
            builder.AppendLine();
            builder.AppendLine("Operations:");

            foreach (var operation in result.Operations)
            {
                builder.AppendLine($"{operation.Status}:");
                builder.AppendLine($"  From: {operation.SourcePath}");
                builder.AppendLine($"  To:   {operation.DestinationPath}");

                if (!string.IsNullOrWhiteSpace(operation.SkipReason))
                {
                    builder.AppendLine($"  Reason: {operation.SkipReason}");
                }

                if (!string.IsNullOrWhiteSpace(operation.StatusMessage))
                {
                    builder.AppendLine($"  Message: {operation.StatusMessage}");
                }

                if (!string.IsNullOrWhiteSpace(operation.BackupPath))
                {
                    builder.AppendLine($"  Backup: {operation.BackupPath}");
                }

                builder.AppendLine($"  Tags: {string.Join(", ", operation.Tags)}");
                builder.AppendLine();
            }
        }

        if (result.Conversions.Count > 0)
        {
            builder.AppendLine();
            builder.AppendLine("Conversions:");

            foreach (var conversion in result.Conversions)
            {
                builder.AppendLine($"{conversion.Status}:");
                builder.AppendLine($"  From: {conversion.SourcePath}");
                builder.AppendLine($"  To:   {conversion.OutputPath}");
                builder.AppendLine($"  Target: {conversion.TargetFormat}");
                builder.AppendLine($"  Tool:   {conversion.Converter}");
                if (!string.IsNullOrWhiteSpace(conversion.Message))
                {
                    builder.AppendLine($"  Message: {conversion.Message}");
                }

                builder.AppendLine();
            }
        }

        if (result.CleanedDirectories.Count > 0)
        {
            builder.AppendLine("Removed empty folders:");
            foreach (var directory in result.CleanedDirectories)
            {
                builder.AppendLine($"  {directory}");
            }
        }

        return builder.ToString().TrimEnd();
    }

    private static List<string> GetCandidateFiles(
        string sourceRoot,
        string destinationRoot,
        RulesConfig rules,
        bool topLevelOnly,
        IReadOnlyCollection<string> includeExtensions,
        IProgress<OrganizerProgressInfo>? progress,
        CancellationToken cancellationToken)
    {
        progress?.Report(new OrganizerProgressInfo("Scanning", 0, 0, "Scanning source folder...", true));

        var candidates = new List<string>();
        var scannedFiles = 0;

        foreach (var path in EnumerateFilesSafe(sourceRoot, destinationRoot, rules.ExcludeRules, topLevelOnly, progress, cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (IsSameOrChildPath(path, destinationRoot))
            {
                continue;
            }

            if (IsExcluded(path, sourceRoot, rules.ExcludeRules))
            {
                continue;
            }

            if (includeExtensions.Count > 0)
            {
                var extension = Path.GetExtension(path).ToLowerInvariant();
                if (!includeExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
                {
                    continue;
                }
            }

            candidates.Add(path);
            scannedFiles++;
            progress?.Report(new OrganizerProgressInfo("Scanning", scannedFiles, 0, $"Scanning {Path.GetFileName(path)}", true));
        }

        return candidates
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static List<FileActionRecord> BuildPlan(
        IReadOnlyList<string> candidateFiles,
        string destinationRoot,
        RulesConfig rules,
        OrganizerOptions options,
        IProgress<OrganizerProgressInfo>? progress,
        CancellationToken cancellationToken)
    {
        var categoryLookup = rules.CategoryRules
            .SelectMany(rule => rule.Extensions.Select(extension => new KeyValuePair<string, string>(extension, rule.Name)))
            .ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.OrdinalIgnoreCase);

        var plan = new List<FileActionRecord>();
        var plannedDestinations = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        for (var index = 0; index < candidateFiles.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                var file = new FileInfo(candidateFiles[index]);
                var action = CreateBaseAction(file, destinationRoot, rules, categoryLookup, options);
                var resolved = ResolveCollision(action, file, plan, plannedDestinations, options.DuplicateHandling);
                plan.Add(resolved);

                if (resolved.ShouldExecute)
                {
                    plannedDestinations[resolved.DestinationPath] = plan.Count - 1;
                }

                progress?.Report(new OrganizerProgressInfo("Planning", index + 1, candidateFiles.Count, $"Planning {file.Name}"));
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Security.SecurityException)
            {
                var skipped = CreateUnavailableAction(candidateFiles[index], options.DuplicateHandling, $"Skipped because the file could not be inspected: {ex.Message}");
                plan.Add(skipped);
                progress?.Report(new OrganizerProgressInfo("Planning", index + 1, candidateFiles.Count, $"Skipping {Path.GetFileName(candidateFiles[index])}"));
            }
        }

        return plan;
    }

    private static FileActionRecord CreateBaseAction(
        FileInfo file,
        string destinationRoot,
        RulesConfig rules,
        IReadOnlyDictionary<string, string> categoryLookup,
        OrganizerOptions options)
    {
        var extension = file.Extension.ToLowerInvariant();
        var category = categoryLookup.TryGetValue(extension, out var foundCategory) ? foundCategory : "Other";
        var keywordGroup = GetKeywordGroup(file.Name, rules.KeywordRules);
        var sizeBand = GetSizeBand(file.Length, rules.SizeBands);
        var destinationFolder = GetDestinationFolder(destinationRoot, options.OrganizationMode, category, keywordGroup, sizeBand, file.LastWriteTime, extension, options.UseExtensionSubfolders);
        var destinationName = BuildDestinationFileName(file, category, keywordGroup, sizeBand, rules.NamingRules, options);
        var destinationPath = Path.Combine(destinationFolder, destinationName);
        var tags = BuildTags(file, category, keywordGroup, sizeBand, options.OrganizationMode);
        var groupName = GetGroupName(options.OrganizationMode, category, keywordGroup, sizeBand, file.LastWriteTime, extension);

        return new FileActionRecord
        {
            SourcePath = file.FullName,
            DestinationPath = destinationPath,
            DestinationFolder = destinationFolder,
            GroupName = groupName,
            Category = category,
            Extension = extension,
            SizeBytes = file.Length,
            LastModified = file.LastWriteTime.ToString("o"),
            Tags = tags,
            DuplicateHandling = options.DuplicateHandling,
            KeywordGroup = keywordGroup,
            Status = FileOperationStatus.Planned
        };
    }

    private static FileActionRecord ResolveCollision(
        FileActionRecord action,
        FileInfo sourceFile,
        List<FileActionRecord> plan,
        Dictionary<string, int> plannedDestinations,
        DuplicateHandlingMode duplicateHandling)
    {
        if (plannedDestinations.TryGetValue(action.DestinationPath, out var existingPlanIndex))
        {
            return ResolvePlannedCollision(action, sourceFile, plan, existingPlanIndex, duplicateHandling);
        }

        if (!File.Exists(action.DestinationPath))
        {
            return action;
        }

        return duplicateHandling switch
        {
            DuplicateHandlingMode.Rename => action with
            {
                DestinationPath = GetUniqueDestinationPath(action.DestinationPath, plan.Select(item => item.DestinationPath))
            },
            DuplicateHandlingMode.Skip => action with
            {
                ShouldExecute = false,
                DestinationOriginallyExisted = true,
                ExistingDestinationPath = action.DestinationPath,
                SkipReason = "Skipped because the destination file already exists.",
                Status = FileOperationStatus.Skipped
            },
            DuplicateHandlingMode.Overwrite => action with
            {
                DestinationOriginallyExisted = true,
                ExistingDestinationPath = action.DestinationPath
            },
            DuplicateHandlingMode.KeepNewest => sourceFile.LastWriteTimeUtc > File.GetLastWriteTimeUtc(action.DestinationPath)
                ? action with
                {
                    DestinationOriginallyExisted = true,
                    ExistingDestinationPath = action.DestinationPath
                }
                : action with
                {
                    ShouldExecute = false,
                    DestinationOriginallyExisted = true,
                    ExistingDestinationPath = action.DestinationPath,
                    SkipReason = "Skipped because the existing destination file is newer.",
                    Status = FileOperationStatus.Skipped
                },
            DuplicateHandlingMode.CompareByHash => FilesHaveSameHash(sourceFile.FullName, action.DestinationPath)
                ? action with
                {
                    ShouldExecute = false,
                    DestinationOriginallyExisted = true,
                    ExistingDestinationPath = action.DestinationPath,
                    SkipReason = "Skipped because the destination file already has matching content.",
                    Status = FileOperationStatus.Skipped
                }
                : action with
                {
                    DestinationPath = GetUniqueDestinationPath(action.DestinationPath, plan.Select(item => item.DestinationPath))
                },
            _ => action
        };
    }

    private static FileActionRecord ResolvePlannedCollision(
        FileActionRecord action,
        FileInfo sourceFile,
        List<FileActionRecord> plan,
        int existingPlanIndex,
        DuplicateHandlingMode duplicateHandling)
    {
        var existing = plan[existingPlanIndex];

        switch (duplicateHandling)
        {
            case DuplicateHandlingMode.Rename:
                return action with
                {
                    DestinationPath = GetUniqueDestinationPath(action.DestinationPath, plan.Select(item => item.DestinationPath))
                };
            case DuplicateHandlingMode.Skip:
                return action with
                {
                    ShouldExecute = false,
                    ExistingDestinationPath = existing.DestinationPath,
                    SkipReason = "Skipped because another file in this run already targeted the same destination.",
                    Status = FileOperationStatus.Skipped
                };
            case DuplicateHandlingMode.Overwrite:
                plan[existingPlanIndex] = existing with
                {
                    ShouldExecute = false,
                    SkipReason = "Superseded by a later file because duplicate handling is set to Overwrite.",
                    Status = FileOperationStatus.Skipped
                };
                return action;
            case DuplicateHandlingMode.KeepNewest:
                if (sourceFile.LastWriteTimeUtc > DateTime.Parse(existing.LastModified).ToUniversalTime())
                {
                    plan[existingPlanIndex] = existing with
                    {
                        ShouldExecute = false,
                        SkipReason = "Superseded by a newer file in this run.",
                        Status = FileOperationStatus.Skipped
                    };
                    return action;
                }

                return action with
                {
                    ShouldExecute = false,
                    ExistingDestinationPath = existing.DestinationPath,
                    SkipReason = "Skipped because another planned file for the same destination is newer.",
                    Status = FileOperationStatus.Skipped
                };
            case DuplicateHandlingMode.CompareByHash:
                if (FilesHaveSameHash(sourceFile.FullName, existing.SourcePath))
                {
                    return action with
                    {
                        ShouldExecute = false,
                        ExistingDestinationPath = existing.DestinationPath,
                        SkipReason = "Skipped because another file in this run has matching content for the same destination.",
                        Status = FileOperationStatus.Skipped
                    };
                }

                return action with
                {
                    DestinationPath = GetUniqueDestinationPath(action.DestinationPath, plan.Select(item => item.DestinationPath))
                };
            default:
                return action;
        }
    }

    private static List<FileActionRecord> ExecutePlan(
        IReadOnlyList<FileActionRecord> plan,
        OrganizerOptions options,
        string runDirectory,
        IProgress<OrganizerProgressInfo>? progress,
        CancellationToken cancellationToken)
    {
        var executed = new List<FileActionRecord>();
        for (var index = 0; index < plan.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var item = plan[index];
            if (!item.ShouldExecute || options.WhatIf)
            {
                executed.Add(item with
                {
                    Status = item.Status == FileOperationStatus.Skipped ? FileOperationStatus.Skipped : FileOperationStatus.Planned
                });
                progress?.Report(new OrganizerProgressInfo("Executing", index + 1, plan.Count, $"Skipping {Path.GetFileName(item.SourcePath)}"));
                continue;
            }

            var backupPath = string.Empty;
            try
            {
                Directory.CreateDirectory(item.DestinationFolder);

                if (item.DestinationOriginallyExisted && File.Exists(item.DestinationPath))
                {
                    backupPath = BackupExistingDestination(item.DestinationPath, runDirectory);
                }

                if (options.Mode == OperationMode.Copy)
                {
                    File.Copy(item.SourcePath, item.DestinationPath, overwrite: item.DestinationOriginallyExisted);
                }
                else
                {
                    if (item.DestinationOriginallyExisted && File.Exists(item.DestinationPath))
                    {
                        File.Delete(item.DestinationPath);
                    }

                    File.Move(item.SourcePath, item.DestinationPath);
                }

                executed.Add(item with
                {
                    BackupPath = string.IsNullOrWhiteSpace(backupPath) ? null : backupPath,
                    Status = FileOperationStatus.Executed
                });

                progress?.Report(new OrganizerProgressInfo("Executing", index + 1, plan.Count, $"Processed {Path.GetFileName(item.SourcePath)}"));
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Security.SecurityException)
            {
                TryRestoreFailedBackup(item.DestinationPath, backupPath);
                executed.Add(item with
                {
                    BackupPath = string.IsNullOrWhiteSpace(backupPath) ? null : backupPath,
                    Status = FileOperationStatus.Failed,
                    StatusMessage = ex.Message
                });
                progress?.Report(new OrganizerProgressInfo("Executing", index + 1, plan.Count, $"Failed {Path.GetFileName(item.SourcePath)}"));
            }
        }

        return executed;
    }

    private static List<string> CleanupEmptyDirectories(
        string sourceRoot,
        string destinationRoot,
        IProgress<OrganizerProgressInfo>? progress,
        CancellationToken cancellationToken)
    {
        var directories = Directory.EnumerateDirectories(sourceRoot, "*", SearchOption.AllDirectories)
            .Where(path => !IsSameOrChildPath(path, destinationRoot))
            .OrderByDescending(path => path.Length)
            .ToList();

        var removed = new List<string>();
        for (var index = 0; index < directories.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var directory = directories[index];
            if (!Directory.Exists(directory))
            {
                continue;
            }

            if (Directory.EnumerateFileSystemEntries(directory).Any())
            {
                continue;
            }

            Directory.Delete(directory, false);
            removed.Add(directory);
            progress?.Report(new OrganizerProgressInfo("Cleanup", index + 1, directories.Count, $"Removed empty folder {directory}"));
        }

        return removed;
    }

    private static void WriteTagManifest(IEnumerable<FileActionRecord> operations, string manifestPath)
    {
        var directory = Path.GetDirectoryName(manifestPath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var payload = operations.Select(item => new
        {
            item.SourcePath,
            item.DestinationPath,
            item.Category,
            item.Extension,
            item.SizeBytes,
            item.LastModified,
            item.Tags
        });

        File.WriteAllText(manifestPath, JsonSerializer.Serialize(payload, JsonOptions));
    }

    private static void WriteLogArtifacts(OrganizerRunResult result)
    {
        var document = new OperationLogDocument
        {
            RunId = result.RunId,
            RunKind = result.RunKind,
            Timestamp = result.Timestamp,
            SourceRoot = result.SourceRoot,
            DestinationRoot = result.DestinationRoot,
            Mode = result.Mode,
            DuplicateHandling = result.DuplicateHandling,
            OrganizationMode = result.OrganizationMode,
            PdfRenameMode = result.PdfRenameMode,
            WhatIf = result.WhatIf,
            TagManifestPath = result.TagManifestPath,
            IncludeConvert = result.IncludeConvert,
            ConvertSourceFormat = result.ConvertSourceFormat,
            ConvertTargetFormat = result.ConvertTargetFormat,
            ConvertQuality = result.ConvertQuality,
            Operations = result.Operations,
            Conversions = result.Conversions,
            CleanedDirectories = result.CleanedDirectories
        };

        File.WriteAllText(result.LogFilePath, result.DetailedReport);
        File.WriteAllText(result.OperationLogPath, JsonSerializer.Serialize(document, JsonOptions));
    }

    private static void AppendHistoryEntry(OrganizerRunResult result)
    {
        AppStorage.AppendHistory(new RunHistoryEntry
        {
            RunId = result.RunId,
            RunKind = result.RunKind,
            Timestamp = result.Timestamp,
            SourceRoot = result.SourceRoot,
            DestinationRoot = result.DestinationRoot,
            Mode = result.Mode,
            DuplicateHandling = result.DuplicateHandling,
            OrganizationMode = result.OrganizationMode,
            PdfRenameMode = result.PdfRenameMode,
            WhatIf = result.WhatIf,
            PlannedCount = result.Operations.Count,
            ExecutedCount = result.Operations.Count(item => item.Status == FileOperationStatus.Executed),
            SkippedCount = result.Operations.Count(item => item.Status == FileOperationStatus.Skipped),
            ConversionCount = result.Conversions.Count(item => item.Status == ConversionStatus.Executed),
            RulesPath = result.RulesPath,
            TagManifestPath = result.TagManifestPath,
            LogFilePath = result.LogFilePath,
            OperationLogPath = result.OperationLogPath,
            SummaryText = result.SummaryText,
            CanUndo = !result.WhatIf && result.Operations.Any(item => item.Status == FileOperationStatus.Executed),
            IsUndone = false
        });
    }

    private static bool IsExcluded(string path, string sourceRoot, ExcludeRules rules)
    {
        try
        {
            var attributes = File.GetAttributes(path);
            if (rules.IgnoreHiddenFiles && attributes.HasFlag(FileAttributes.Hidden))
            {
                return true;
            }

            if (rules.IgnoreSystemFiles && attributes.HasFlag(FileAttributes.System))
            {
                return true;
            }

            var extension = Path.GetExtension(path).ToLowerInvariant();
            if (rules.Extensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
            {
                return true;
            }

            var relativePath = Path.GetRelativePath(sourceRoot, path);
            var parts = relativePath.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            foreach (var part in parts.Take(Math.Max(parts.Length - 1, 0)))
            {
                if (rules.FolderNames.Contains(part, StringComparer.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            var fileInfo = new FileInfo(path);
            if (rules.MinFileSizeBytes.HasValue && fileInfo.Length < rules.MinFileSizeBytes.Value)
            {
                return true;
            }

            if (rules.MaxFileSizeBytes.HasValue && fileInfo.Length > rules.MaxFileSizeBytes.Value)
            {
                return true;
            }

            return false;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Security.SecurityException)
        {
            return true;
        }
    }

    private static IEnumerable<string> EnumerateFilesSafe(
        string sourceRoot,
        string destinationRoot,
        ExcludeRules rules,
        bool topLevelOnly,
        IProgress<OrganizerProgressInfo>? progress,
        CancellationToken cancellationToken)
    {
        var pending = new Stack<string>();
        pending.Push(sourceRoot);

        while (pending.Count > 0)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var currentDirectory = pending.Pop();

            foreach (var file in SafeEnumerateFiles(currentDirectory, progress))
            {
                yield return file;
            }

            if (topLevelOnly)
            {
                continue;
            }

            var childDirectories = SafeEnumerateDirectories(currentDirectory, progress)
                .OrderByDescending(path => path, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var directory in childDirectories)
            {
                if (ShouldSkipDirectory(directory, sourceRoot, destinationRoot, rules))
                {
                    continue;
                }

                pending.Push(directory);
            }
        }
    }

    private static IEnumerable<string> SafeEnumerateFiles(string directoryPath, IProgress<OrganizerProgressInfo>? progress)
    {
        try
        {
            return Directory.EnumerateFiles(directoryPath, "*", SearchOption.TopDirectoryOnly).ToList();
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Security.SecurityException)
        {
            progress?.Report(new OrganizerProgressInfo("Scanning", 0, 0, $"Skipped unreadable folder {directoryPath}: {ex.Message}", true));
            return Array.Empty<string>();
        }
    }

    private static IEnumerable<string> SafeEnumerateDirectories(string directoryPath, IProgress<OrganizerProgressInfo>? progress)
    {
        try
        {
            return Directory.EnumerateDirectories(directoryPath, "*", SearchOption.TopDirectoryOnly).ToList();
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Security.SecurityException)
        {
            progress?.Report(new OrganizerProgressInfo("Scanning", 0, 0, $"Skipped unreadable folder {directoryPath}: {ex.Message}", true));
            return Array.Empty<string>();
        }
    }

    private static bool ShouldSkipDirectory(string directoryPath, string sourceRoot, string destinationRoot, ExcludeRules rules)
    {
        if (IsSameOrChildPath(directoryPath, destinationRoot))
        {
            return true;
        }

        try
        {
            var attributes = File.GetAttributes(directoryPath);
            if (attributes.HasFlag(FileAttributes.ReparsePoint))
            {
                return true;
            }

            if (rules.IgnoreHiddenFiles && attributes.HasFlag(FileAttributes.Hidden))
            {
                return true;
            }

            if (rules.IgnoreSystemFiles && attributes.HasFlag(FileAttributes.System))
            {
                return true;
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Security.SecurityException)
        {
            return true;
        }

        var relativePath = Path.GetRelativePath(sourceRoot, directoryPath);
        var parts = relativePath.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries);
        return parts.Any(part => rules.FolderNames.Contains(part, StringComparer.OrdinalIgnoreCase));
    }

    private static FileActionRecord CreateUnavailableAction(string path, DuplicateHandlingMode duplicateHandling, string reason)
    {
        return new FileActionRecord
        {
            SourcePath = path,
            DestinationPath = path,
            DestinationFolder = Path.GetDirectoryName(path) ?? string.Empty,
            GroupName = "Unavailable",
            Category = "Unavailable",
            Extension = Path.GetExtension(path).ToLowerInvariant(),
            SizeBytes = 0,
            LastModified = DateTimeOffset.Now.ToString("o"),
            Tags = Array.Empty<string>(),
            DuplicateHandling = duplicateHandling,
            ShouldExecute = false,
            SkipReason = reason,
            Status = FileOperationStatus.Skipped
        };
    }

    private static void TryRestoreFailedBackup(string destinationPath, string backupPath)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(backupPath) || !File.Exists(backupPath) || File.Exists(destinationPath))
            {
                return;
            }

            var directory = Path.GetDirectoryName(destinationPath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.Copy(backupPath, destinationPath, overwrite: true);
        }
        catch
        {
            // Best-effort restore only.
        }
    }

    private static string GetDestinationFolder(
        string destinationRoot,
        OrganizationMode organizationMode,
        string category,
        string keywordGroup,
        string sizeBand,
        DateTime modifiedDate,
        string extension,
        bool useExtensionSubfolders)
    {
        var segments = organizationMode switch
        {
            OrganizationMode.ByType => new List<string> { category },
            OrganizationMode.ByExtension => new List<string> { string.IsNullOrWhiteSpace(extension) ? "no-extension" : extension.TrimStart('.').ToLowerInvariant() },
            OrganizationMode.ByYearMonth => new List<string> { modifiedDate.ToString("yyyy"), modifiedDate.ToString("MM") },
            OrganizationMode.ByDateModified => new List<string> { modifiedDate.ToString("yyyy"), modifiedDate.ToString("MM"), modifiedDate.ToString("dd") },
            OrganizationMode.BySize => new List<string> { sizeBand },
            OrganizationMode.ByKeyword => new List<string> { keywordGroup },
            OrganizationMode.CategoryYearMonth => new List<string> { category, modifiedDate.ToString("yyyy"), modifiedDate.ToString("MM") },
            _ => new List<string> { category }
        };

        if (useExtensionSubfolders &&
            (organizationMode == OrganizationMode.ByType || organizationMode == OrganizationMode.CategoryYearMonth) &&
            !string.IsNullOrWhiteSpace(extension))
        {
            segments.Add(extension.TrimStart('.').ToLowerInvariant());
        }

        var path = destinationRoot;
        foreach (var segment in segments.Where(segment => !string.IsNullOrWhiteSpace(segment)))
        {
            path = Path.Combine(path, segment);
        }

        return path;
    }

    private static string BuildDestinationFileName(FileInfo file, string category, string keywordGroup, string sizeBand, NamingRules namingRules, OrganizerOptions options)
    {
        var template = namingRules.FileNameTemplate ?? "{originalName}";
        var baseName = ResolveDestinationBaseName(file, namingRules, options);
        var usesExtensionToken = template.Contains("{extension}", StringComparison.OrdinalIgnoreCase);

        var rendered = template
            .Replace("{originalName}", baseName, StringComparison.OrdinalIgnoreCase)
            .Replace("{category}", category, StringComparison.OrdinalIgnoreCase)
            .Replace("{keywordGroup}", keywordGroup, StringComparison.OrdinalIgnoreCase)
            .Replace("{sizeBand}", sizeBand, StringComparison.OrdinalIgnoreCase)
            .Replace("{yyyy}", file.LastWriteTime.ToString("yyyy"), StringComparison.OrdinalIgnoreCase)
            .Replace("{MM}", file.LastWriteTime.ToString("MM"), StringComparison.OrdinalIgnoreCase)
            .Replace("{dd}", file.LastWriteTime.ToString("dd"), StringComparison.OrdinalIgnoreCase)
            .Replace("{extension}", file.Extension.TrimStart('.').ToLowerInvariant(), StringComparison.OrdinalIgnoreCase);

        var sanitized = SanitizeFileName(rendered, namingRules.InvalidCharacterReplacement);
        if (string.IsNullOrWhiteSpace(sanitized))
        {
            sanitized = "file";
        }

        return usesExtensionToken || string.IsNullOrWhiteSpace(file.Extension)
            ? sanitized
            : sanitized + file.Extension;
    }

    private static string SanitizeFileName(string name, string replacement)
    {
        var result = name;
        foreach (var invalidChar in Path.GetInvalidFileNameChars())
        {
            result = result.Replace(invalidChar.ToString(), replacement, StringComparison.Ordinal);
        }

        return result.Trim().TrimEnd('.');
    }

    private static string ResolveDestinationBaseName(FileInfo file, NamingRules namingRules, OrganizerOptions options)
    {
        var originalName = Path.GetFileNameWithoutExtension(file.Name);
        if (options.PdfRenameMode == PdfRenameMode.Disabled)
        {
            return originalName;
        }

        var extension = file.Extension.ToLowerInvariant();
        if (IsImageExtension(extension))
        {
            return options.MediaImagesUseMetadata
                ? MediaMetadataService.ResolveImageBaseName(file, namingRules.InvalidCharacterReplacement)
                : originalName;
        }

        if (IsAudioExtension(extension) && !options.MediaMusicUseMetadata)
        {
            return originalName;
        }

        if (IsVideoExtension(extension) && !options.MediaVideoUseMetadata)
        {
            return originalName;
        }

        return MetadataNamingService.ResolveBaseName(file, options.PdfRenameMode, namingRules.InvalidCharacterReplacement);
    }

    private static bool IsImageExtension(string extension)
    {
        return extension is ".jpg" or ".jpeg" or ".png" or ".bmp" or ".gif" or ".tif" or ".tiff" or ".heic" or ".webp";
    }

    private static bool IsAudioExtension(string extension)
    {
        return extension is ".mp3" or ".wav" or ".flac" or ".aac" or ".ogg" or ".m4a" or ".wma" or ".aiff" or ".alac";
    }

    private static bool IsVideoExtension(string extension)
    {
        return extension is ".mp4" or ".mov" or ".avi" or ".mkv" or ".wmv" or ".webm" or ".m4v" or ".mpeg" or ".mpg";
    }

    private static string GetKeywordGroup(string fileName, IReadOnlyList<KeywordRule> keywordRules)
    {
        var baseName = Path.GetFileNameWithoutExtension(fileName);
        foreach (var keywordRule in keywordRules)
        {
            if (keywordRule.Keywords.Any(keyword => baseName.Contains(keyword, StringComparison.OrdinalIgnoreCase)))
            {
                return string.IsNullOrWhiteSpace(keywordRule.Name) ? "Other" : keywordRule.Name;
            }
        }

        return "Other";
    }

    private static string GetSizeBand(long sizeBytes, IReadOnlyList<SizeBandRule> sizeBands)
    {
        foreach (var band in sizeBands)
        {
            if (!band.MaxBytes.HasValue || sizeBytes <= band.MaxBytes.Value)
            {
                return string.IsNullOrWhiteSpace(band.Name) ? "Other" : band.Name;
            }
        }

        return "Other";
    }

    private static string GetGroupName(OrganizationMode mode, string category, string keywordGroup, string sizeBand, DateTime modifiedDate, string extension)
    {
        return mode switch
        {
            OrganizationMode.ByType => category,
            OrganizationMode.ByExtension => string.IsNullOrWhiteSpace(extension) ? "no-extension" : extension.TrimStart('.').ToLowerInvariant(),
            OrganizationMode.ByYearMonth => modifiedDate.ToString("yyyy-MM"),
            OrganizationMode.ByDateModified => modifiedDate.ToString("yyyy-MM-dd"),
            OrganizationMode.BySize => sizeBand,
            OrganizationMode.ByKeyword => keywordGroup,
            OrganizationMode.CategoryYearMonth => category,
            _ => category
        };
    }

    private static string[] BuildTags(FileInfo file, string category, string keywordGroup, string sizeBand, OrganizationMode organizationMode)
    {
        var tags = new List<string>
        {
            $"category:{category.ToLowerInvariant()}",
            string.IsNullOrWhiteSpace(file.Extension) ? "ext:none" : $"ext:{file.Extension.TrimStart('.').ToLowerInvariant()}",
            $"year:{file.LastWriteTime:yyyy}",
            $"month:{file.LastWriteTime:MM}",
            $"size:{sizeBand.ToLowerInvariant()}",
            $"mode:{organizationMode}"
        };

        if (!string.Equals(keywordGroup, "Other", StringComparison.OrdinalIgnoreCase))
        {
            tags.Add($"keyword:{keywordGroup.ToLowerInvariant()}");
        }

        return tags.ToArray();
    }

    private static string BackupExistingDestination(string destinationPath, string runDirectory)
    {
        var backupDirectory = Path.Combine(runDirectory, "backups");
        Directory.CreateDirectory(backupDirectory);

        var backupPath = Path.Combine(backupDirectory, $"{Guid.NewGuid():N}{Path.GetExtension(destinationPath)}");
        File.Copy(destinationPath, backupPath, overwrite: true);
        return backupPath;
    }

    private static string GetUniqueDestinationPath(string desiredPath, IEnumerable<string> reservedPaths)
    {
        var reserved = new HashSet<string>(reservedPaths, StringComparer.OrdinalIgnoreCase);
        if (!File.Exists(desiredPath) && !reserved.Contains(desiredPath))
        {
            return desiredPath;
        }

        var directory = Path.GetDirectoryName(desiredPath) ?? throw new InvalidOperationException($"Unable to determine the destination directory for {desiredPath}");
        var fileName = Path.GetFileNameWithoutExtension(desiredPath);
        var extension = Path.GetExtension(desiredPath);
        var counter = 1;

        while (true)
        {
            var candidate = Path.Combine(directory, $"{fileName} ({counter}){extension}");
            if (!File.Exists(candidate) && !reserved.Contains(candidate))
            {
                return candidate;
            }

            counter++;
        }
    }

    private static bool FilesHaveSameHash(string leftPath, string rightPath)
    {
        using var leftStream = File.OpenRead(leftPath);
        using var rightStream = File.OpenRead(rightPath);
        using var hashAlgorithm = SHA256.Create();

        var leftHash = hashAlgorithm.ComputeHash(leftStream);
        var rightHash = hashAlgorithm.ComputeHash(rightStream);
        return leftHash.SequenceEqual(rightHash);
    }

    private static bool IsSameOrChildPath(string path, string parentPath)
    {
        var normalizedPath = Path.TrimEndingDirectorySeparator(Path.GetFullPath(path));
        var normalizedParent = Path.TrimEndingDirectorySeparator(Path.GetFullPath(parentPath));

        if (string.Equals(normalizedPath, normalizedParent, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return normalizedPath.StartsWith(normalizedParent + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
    }

    private static bool PathsEqual(string left, string right)
    {
        return string.Equals(
            Path.TrimEndingDirectorySeparator(Path.GetFullPath(left)),
            Path.TrimEndingDirectorySeparator(Path.GetFullPath(right)),
            StringComparison.OrdinalIgnoreCase);
    }
}
