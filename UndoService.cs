using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading;

namespace FileOrganizer;

internal static class UndoService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    public static UndoResult UndoLastRun(IProgress<OrganizerProgressInfo>? progress = null, CancellationToken cancellationToken = default)
    {
        var history = AppStorage.LoadHistory();
        var target = history.FirstOrDefault(entry => entry.RunKind == RunKind.Organize && !entry.WhatIf && entry.CanUndo && !entry.IsUndone);
        if (target is null)
        {
            throw new InvalidOperationException("There is no completed run available to undo.");
        }

        return UndoRun(target.RunId, progress, cancellationToken);
    }

    public static UndoResult UndoRun(string runId, IProgress<OrganizerProgressInfo>? progress = null, CancellationToken cancellationToken = default)
    {
        var history = AppStorage.LoadHistory();
        var target = history.FirstOrDefault(entry => string.Equals(entry.RunId, runId, StringComparison.OrdinalIgnoreCase));
        if (target is null)
        {
            throw new InvalidOperationException($"History entry not found: {runId}");
        }

        if (target.WhatIf || !target.CanUndo || target.IsUndone)
        {
            throw new InvalidOperationException("The selected history entry cannot be undone.");
        }

        if (!File.Exists(target.OperationLogPath))
        {
            throw new FileNotFoundException("The operation log for the selected run is missing.", target.OperationLogPath);
        }

        var operationLog = JsonSerializer.Deserialize<OperationLogDocument>(File.ReadAllText(target.OperationLogPath), JsonOptions)
            ?? throw new InvalidOperationException("Unable to read the operation log for the selected run.");

        var appPaths = AppStorage.GetPaths();
        var undoRunId = $"undo-{DateTimeOffset.Now:yyyyMMdd-HHmmss}-{Guid.NewGuid():N}";
        var undoDirectory = Path.Combine(appPaths.LogsDirectory, undoRunId);
        Directory.CreateDirectory(undoDirectory);
        var logPath = Path.Combine(undoDirectory, $"{undoRunId}.log");

        var messages = new List<string>();
        var executableOperations = operationLog.Operations.Where(item => item.Status == FileOperationStatus.Executed).Reverse().ToList();

        for (var index = 0; index < executableOperations.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var item = executableOperations[index];

            progress?.Report(new OrganizerProgressInfo("Undo", index + 1, executableOperations.Count, $"Undoing {Path.GetFileName(item.DestinationPath)}"));

            RemoveGeneratedOutputs(item, messages);

            if (operationLog.Mode == OperationMode.Move)
            {
                UndoMove(item, messages);
            }
            else
            {
                UndoCopy(item, messages);
            }
        }

        RemoveTagManifest(operationLog.TagManifestPath, messages);
        RemoveEmptyDestinationDirectories(operationLog.DestinationRoot, messages, cancellationToken);

        var builder = new StringBuilder();
        builder.AppendLine($"Undo run ID: {undoRunId}");
        builder.AppendLine($"Target run ID: {target.RunId}");
        builder.AppendLine($"Target log: {target.LogFilePath}");
        builder.AppendLine($"Log file: {logPath}");
        builder.AppendLine();
        builder.AppendLine("Undo actions:");
        foreach (var message in messages)
        {
            builder.AppendLine($" - {message}");
        }

        var detailedReport = builder.ToString().TrimEnd();
        File.WriteAllText(logPath, detailedReport);

        AppStorage.UpdateHistory(entries =>
        {
            var updated = new List<RunHistoryEntry>();
            foreach (var entry in entries)
            {
                if (string.Equals(entry.RunId, target.RunId, StringComparison.OrdinalIgnoreCase))
                {
                    updated.Add(entry with
                    {
                        IsUndone = true,
                        UndoneByRunId = undoRunId,
                        CanUndo = false
                    });
                }
                else
                {
                    updated.Add(entry);
                }
            }

            updated.Insert(0, new RunHistoryEntry
            {
                RunId = undoRunId,
                RunKind = RunKind.Undo,
                Timestamp = DateTimeOffset.Now,
                SourceRoot = target.SourceRoot,
                DestinationRoot = target.DestinationRoot,
                Mode = target.Mode,
                DuplicateHandling = target.DuplicateHandling,
                OrganizationMode = target.OrganizationMode,
                PdfRenameMode = target.PdfRenameMode,
                WhatIf = false,
                PlannedCount = executableOperations.Count,
                ExecutedCount = executableOperations.Count,
                SkippedCount = 0,
                ConversionCount = target.ConversionCount,
                RulesPath = target.RulesPath,
                TagManifestPath = target.TagManifestPath,
                LogFilePath = logPath,
                OperationLogPath = target.OperationLogPath,
                SummaryText = $"Undid run {target.RunId}",
                CanUndo = false,
                IsUndone = false
            });

            return updated;
        });

        return new UndoResult
        {
            RunId = undoRunId,
            TargetRunId = target.RunId,
            LogFilePath = logPath,
            SummaryText = $"Undid run {target.RunId}",
            DetailedReport = detailedReport
        };
    }

    private static void UndoMove(FileActionRecord item, List<string> messages)
    {
        if (!File.Exists(item.DestinationPath))
        {
            messages.Add($"Skipped move undo for missing destination: {item.DestinationPath}");
        }
        else
        {
            Directory.CreateDirectory(Path.GetDirectoryName(item.SourcePath) ?? throw new InvalidOperationException("Invalid source path."));
            var restorePath = item.SourcePath;
            if (File.Exists(restorePath))
            {
                restorePath = GetUniqueUndoPath(restorePath);
                messages.Add($"Source path already existed. Restored moved file to {restorePath}");
            }
            else
            {
                messages.Add($"Moved {item.DestinationPath} back to {restorePath}");
            }

            File.Move(item.DestinationPath, restorePath);
        }

        if (!string.IsNullOrWhiteSpace(item.BackupPath) && File.Exists(item.BackupPath))
        {
            Directory.CreateDirectory(Path.GetDirectoryName(item.DestinationPath) ?? throw new InvalidOperationException("Invalid destination path."));
            File.Copy(item.BackupPath, item.DestinationPath, overwrite: true);
            messages.Add($"Restored backed up destination file to {item.DestinationPath}");
        }
    }

    private static void UndoCopy(FileActionRecord item, List<string> messages)
    {
        if (File.Exists(item.DestinationPath))
        {
            File.Delete(item.DestinationPath);
            messages.Add($"Deleted copied file {item.DestinationPath}");
        }
        else
        {
            messages.Add($"Copied file already missing, nothing to delete: {item.DestinationPath}");
        }

        if (!string.IsNullOrWhiteSpace(item.BackupPath) && File.Exists(item.BackupPath))
        {
            Directory.CreateDirectory(Path.GetDirectoryName(item.DestinationPath) ?? throw new InvalidOperationException("Invalid destination path."));
            File.Copy(item.BackupPath, item.DestinationPath, overwrite: true);
            messages.Add($"Restored backed up destination file to {item.DestinationPath}");
        }
    }


    private static void RemoveGeneratedOutputs(FileActionRecord item, List<string> messages)
    {
        foreach (var generatedPath in item.GeneratedPaths)
        {
            if (string.IsNullOrWhiteSpace(generatedPath))
            {
                continue;
            }

            if (File.Exists(generatedPath))
            {
                File.Delete(generatedPath);
                messages.Add($"Deleted generated file {generatedPath}");
            }
        }
    }

    private static string GetUniqueUndoPath(string path)
    {
        var directory = Path.GetDirectoryName(path) ?? throw new InvalidOperationException("Invalid path.");
        var fileName = Path.GetFileNameWithoutExtension(path);
        var extension = Path.GetExtension(path);
        var counter = 1;

        while (true)
        {
            var candidate = Path.Combine(directory, $"{fileName} (undo-{counter}){extension}");
            if (!File.Exists(candidate))
            {
                return candidate;
            }

            counter++;
        }
    }

    private static void RemoveTagManifest(string? manifestPath, List<string> messages)
    {
        if (string.IsNullOrWhiteSpace(manifestPath) || !File.Exists(manifestPath))
        {
            return;
        }

        File.Delete(manifestPath);
        messages.Add($"Removed tag manifest {manifestPath}");
    }

    private static void RemoveEmptyDestinationDirectories(string destinationRoot, List<string> messages, CancellationToken cancellationToken)
    {
        if (!Directory.Exists(destinationRoot))
        {
            return;
        }

        var directories = Directory.EnumerateDirectories(destinationRoot, "*", SearchOption.AllDirectories)
            .OrderByDescending(path => path.Length)
            .ToList();

        foreach (var directory in directories)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!Directory.Exists(directory) || Directory.EnumerateFileSystemEntries(directory).Any())
            {
                continue;
            }

            Directory.Delete(directory, false);
            messages.Add($"Removed empty folder {directory}");
        }

        if (Directory.Exists(destinationRoot) && !Directory.EnumerateFileSystemEntries(destinationRoot).Any())
        {
            Directory.Delete(destinationRoot, false);
            messages.Add($"Removed empty folder {destinationRoot}");
        }
    }
}
