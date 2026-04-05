using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace FileOrganizer;

internal sealed record ConversionCapabilities(
    bool ImagePipelineAvailable,
    bool FfmpegAvailable,
    bool LibreOfficeAvailable,
    bool ImageMagickAvailable,
    string AudioVideoPipeline,
    string DocumentPipeline,
    string ImagePipeline,
    string ShellMode);

internal sealed record ConversionBatchResult(
    IReadOnlyList<FileActionRecord> Operations,
    IReadOnlyList<ConversionRecord> Conversions);

internal static class ConversionService
{
    private static readonly HashSet<string> ImageExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff"
    };

    private static readonly HashSet<string> AudioExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".mp3", ".wav", ".flac", ".aac", ".ogg", ".m4a", ".wma", ".aiff", ".alac"
    };

    private static readonly HashSet<string> VideoExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".mp4", ".mov", ".avi", ".mkv", ".wmv", ".webm", ".m4v", ".mpeg", ".mpg"
    };

    private static readonly HashSet<string> DocumentExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".pdf", ".doc", ".docx", ".txt", ".rtf", ".odt", ".md", ".tex", ".ppt", ".pptx", ".odp", ".xls", ".xlsx", ".ods", ".csv", ".tsv"
    };

    public static ConversionCapabilities DetectCapabilities()
    {
        var ffmpegPath = TryFindExecutable("ffmpeg.exe", new[]
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "ffmpeg", "bin", "ffmpeg.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "ffmpeg", "bin", "ffmpeg.exe")
        });

        var sofficePath = TryFindExecutable("soffice.exe", new[]
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "LibreOffice", "program", "soffice.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "LibreOffice", "program", "soffice.exe")
        });

        var magickPath = TryFindExecutable("magick.exe", new[]
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "ImageMagick-7.1.1-Q16-HDRI", "magick.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "ImageMagick-7.1.1-Q16", "magick.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "ImageMagick-7.1.1-Q16", "magick.exe")
        });

        return new ConversionCapabilities(
            ImagePipelineAvailable: true,
            FfmpegAvailable: !string.IsNullOrWhiteSpace(ffmpegPath),
            LibreOfficeAvailable: !string.IsNullOrWhiteSpace(sofficePath),
            ImageMagickAvailable: !string.IsNullOrWhiteSpace(magickPath),
            AudioVideoPipeline: string.IsNullOrWhiteSpace(ffmpegPath) ? "ffmpeg not found" : "ffmpeg ready",
            DocumentPipeline: string.IsNullOrWhiteSpace(sofficePath) ? "LibreOffice not found" : "LibreOffice ready",
            ImagePipeline: !string.IsNullOrWhiteSpace(magickPath) ? "built-in + ImageMagick" : "built-in image pipeline",
            ShellMode: "Browser shell"
        );
    }

    public static ConversionBatchResult Process(
        IReadOnlyList<FileActionRecord> operations,
        OrganizerOptions options,
        IProgress<OrganizerProgressInfo>? progress,
        CancellationToken cancellationToken)
    {
        if (!options.IncludeConvert || string.Equals(options.ConvertTargetFormat, "Keep", StringComparison.OrdinalIgnoreCase))
        {
            return new ConversionBatchResult(operations, Array.Empty<ConversionRecord>());
        }

        var updatedOperations = operations.ToList();
        var conversions = new List<ConversionRecord>();
        var candidates = operations
            .Select((item, index) => new { Item = item, Index = index })
            .Where(entry => options.WhatIf
                ? entry.Item.ShouldExecute && entry.Item.Status != FileOperationStatus.Skipped && entry.Item.Status != FileOperationStatus.Failed
                : entry.Item.Status == FileOperationStatus.Executed)
            .ToList();

        for (var index = 0; index < candidates.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var current = candidates[index];
            progress?.Report(new OrganizerProgressInfo("Convert", index + 1, candidates.Count, $"Converting {Path.GetFileName(current.Item.DestinationPath)}"));

            var conversion = options.WhatIf
                ? BuildPreviewRecord(current.Item, options)
                : ExecuteOne(current.Item, options, cancellationToken);

            conversions.Add(conversion);

            if (conversion.Status == ConversionStatus.Executed)
            {
                var mergedGeneratedPaths = current.Item.GeneratedPaths.Concat(new[] { conversion.OutputPath }).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
                updatedOperations[current.Index] = current.Item with
                {
                    GeneratedPaths = mergedGeneratedPaths,
                    StatusMessage = AppendStatusMessage(current.Item.StatusMessage, $"Converted to {Path.GetFileName(conversion.OutputPath)} via {conversion.Converter}.")
                };
            }
        }

        return new ConversionBatchResult(updatedOperations, conversions);
    }

    private static ConversionRecord BuildPreviewRecord(FileActionRecord item, OrganizerOptions options)
    {
        if (!MatchesSourceFamily(item.Extension, options.ConvertSourceFormat))
        {
            return CreateSkippedRecord(item.DestinationPath, item.DestinationPath, options.ConvertTargetFormat, "preview", "Skipped because the file does not match the selected source format.");
        }

        var supported = ResolveSupportedConverter(item.DestinationPath, options.ConvertTargetFormat);
        if (supported is null)
        {
            return CreateSkippedRecord(item.DestinationPath, item.DestinationPath, options.ConvertTargetFormat, "preview", "Skipped because the requested conversion pair is not supported.");
        }

        var plannedOutput = BuildOutputPath(item.DestinationPath, options.ConvertTargetFormat, ensureUnique: false);
        if (string.Equals(item.DestinationPath, plannedOutput, StringComparison.OrdinalIgnoreCase))
        {
            return CreateSkippedRecord(item.DestinationPath, plannedOutput, options.ConvertTargetFormat, supported.Value.DisplayName, "Skipped because the file is already in the target format.");
        }

        return new ConversionRecord
        {
            SourcePath = item.DestinationPath,
            OutputPath = plannedOutput,
            TargetFormat = options.ConvertTargetFormat,
            Converter = supported.Value.DisplayName,
            Status = ConversionStatus.Planned,
            Message = "Preview only. Conversion would run after the file operation stage."
        };
    }

    private static ConversionRecord ExecuteOne(FileActionRecord item, OrganizerOptions options, CancellationToken cancellationToken)
    {
        if (!MatchesSourceFamily(item.Extension, options.ConvertSourceFormat))
        {
            return CreateSkippedRecord(item.DestinationPath, item.DestinationPath, options.ConvertTargetFormat, "filter", "Skipped because the file does not match the selected source format.");
        }

        var supported = ResolveSupportedConverter(item.DestinationPath, options.ConvertTargetFormat);
        if (supported is null)
        {
            return CreateSkippedRecord(item.DestinationPath, item.DestinationPath, options.ConvertTargetFormat, "none", "Skipped because the requested conversion pair is not supported.");
        }

        var outputPath = BuildOutputPath(item.DestinationPath, options.ConvertTargetFormat, ensureUnique: true);
        if (string.Equals(item.DestinationPath, outputPath, StringComparison.OrdinalIgnoreCase))
        {
            return CreateSkippedRecord(item.DestinationPath, outputPath, options.ConvertTargetFormat, supported.Value.DisplayName, "Skipped because the file is already in the target format.");
        }

        try
        {
            switch (supported.Value.Kind)
            {
                case "image":
                    ConvertImage(item.DestinationPath, outputPath, options.ConvertTargetFormat, options.ConvertQuality);
                    break;
                case "ffmpeg":
                    ConvertWithFfmpeg(item.DestinationPath, outputPath, options.ConvertTargetFormat, options.ConvertQuality, cancellationToken);
                    break;
                case "libreoffice":
                    ConvertWithLibreOffice(item.DestinationPath, outputPath, cancellationToken);
                    break;
                case "magick":
                    ConvertWithImageMagick(item.DestinationPath, outputPath, cancellationToken);
                    break;
                default:
                    return CreateSkippedRecord(item.DestinationPath, outputPath, options.ConvertTargetFormat, supported.Value.DisplayName, "Skipped because no converter is configured for this target.");
            }

            return new ConversionRecord
            {
                SourcePath = item.DestinationPath,
                OutputPath = outputPath,
                TargetFormat = options.ConvertTargetFormat,
                Converter = supported.Value.DisplayName,
                Status = ConversionStatus.Executed,
                Message = "Converted successfully."
            };
        }
        catch (Exception ex)
        {
            TryDeleteFile(outputPath);
            return new ConversionRecord
            {
                SourcePath = item.DestinationPath,
                OutputPath = outputPath,
                TargetFormat = options.ConvertTargetFormat,
                Converter = supported.Value.DisplayName,
                Status = ConversionStatus.Failed,
                Message = ex.Message
            };
        }
    }

    private static (string Kind, string DisplayName)? ResolveSupportedConverter(string sourcePath, string targetFormat)
    {
        var extension = Path.GetExtension(sourcePath);
        var normalizedTarget = NormalizeTarget(targetFormat);
        var capabilities = DetectCapabilities();

        if (normalizedTarget is ".jpg" or ".png" && ImageExtensions.Contains(extension))
        {
            return ("image", "Built-in image pipeline");
        }

        if (normalizedTarget is ".mp3" or ".wav" or ".mp4" && (AudioExtensions.Contains(extension) || VideoExtensions.Contains(extension)))
        {
            return capabilities.FfmpegAvailable ? ("ffmpeg", "ffmpeg") : null;
        }

        if (normalizedTarget == ".pdf")
        {
            if (DocumentExtensions.Contains(extension))
            {
                return capabilities.LibreOfficeAvailable ? ("libreoffice", "LibreOffice") : null;
            }

            if (ImageExtensions.Contains(extension))
            {
                return capabilities.ImageMagickAvailable ? ("magick", "ImageMagick") : null;
            }
        }

        return null;
    }

    private static bool MatchesSourceFamily(string extension, string sourceFormat)
    {
        return sourceFormat.ToLowerInvariant() switch
        {
            "any" => true,
            "document" => DocumentExtensions.Contains(extension),
            "image" => ImageExtensions.Contains(extension),
            "audio" => AudioExtensions.Contains(extension),
            "video" => VideoExtensions.Contains(extension),
            _ => true
        };
    }

    private static string BuildOutputPath(string sourcePath, string targetFormat, bool ensureUnique)
    {
        var normalizedTarget = NormalizeTarget(targetFormat);
        var directory = Path.GetDirectoryName(sourcePath) ?? string.Empty;
        var baseName = Path.GetFileNameWithoutExtension(sourcePath);
        var desired = Path.Combine(directory, baseName + normalizedTarget);

        if (!ensureUnique)
        {
            return desired;
        }

        if (!File.Exists(desired))
        {
            return desired;
        }

        var counter = 1;
        while (true)
        {
            var candidate = Path.Combine(directory, $"{baseName} ({counter}){normalizedTarget}");
            if (!File.Exists(candidate))
            {
                return candidate;
            }

            counter++;
        }
    }

    private static string NormalizeTarget(string targetFormat)
    {
        return targetFormat.ToLowerInvariant() switch
        {
            "pdf" => ".pdf",
            "jpg" => ".jpg",
            "png" => ".png",
            "mp3" => ".mp3",
            "wav" => ".wav",
            "mp4" => ".mp4",
            _ => Path.GetExtension(targetFormat)
        };
    }

    private static void ConvertImage(string sourcePath, string outputPath, string targetFormat, string qualityMode)
    {
        using var sourceImage = Image.FromFile(sourcePath);
        var target = NormalizeTarget(targetFormat);

        if (target == ".png")
        {
            sourceImage.Save(outputPath, ImageFormat.Png);
            return;
        }

        using var bitmap = new Bitmap(sourceImage.Width, sourceImage.Height);
        using (var graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.Black);
            graphics.DrawImage(sourceImage, 0, 0, sourceImage.Width, sourceImage.Height);
        }

        var encoder = ImageCodecInfo.GetImageEncoders().First(codec => codec.FormatID == ImageFormat.Jpeg.Guid);
        using var parameters = new EncoderParameters(1);
        parameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, GetJpegQuality(qualityMode));
        bitmap.Save(outputPath, encoder, parameters);
    }

    private static byte GetJpegQuality(string qualityMode)
    {
        return qualityMode.ToLowerInvariant() switch
        {
            "keep" => 95,
            "compact" => 72,
            _ => 84
        };
    }

    private static void ConvertWithFfmpeg(string sourcePath, string outputPath, string targetFormat, string qualityMode, CancellationToken cancellationToken)
    {
        var ffmpegPath = TryFindExecutable("ffmpeg.exe", Array.Empty<string>())
            ?? throw new InvalidOperationException("ffmpeg was not found on this machine.");

        var target = NormalizeTarget(targetFormat);
        var arguments = target switch
        {
            ".mp3" => $"-y -i {Quote(sourcePath)} -vn -q:a {GetMp3Quality(qualityMode)} {Quote(outputPath)}",
            ".wav" => $"-y -i {Quote(sourcePath)} -vn {Quote(outputPath)}",
            ".mp4" => $"-y -i {Quote(sourcePath)} -movflags +faststart -pix_fmt yuv420p -c:v libx264 -preset medium -crf {GetMp4Crf(qualityMode)} -c:a aac -b:a 160k {Quote(outputPath)}",
            _ => throw new InvalidOperationException($"Unsupported ffmpeg target: {targetFormat}")
        };

        RunProcess(ffmpegPath, arguments, cancellationToken);
    }

    private static string GetMp3Quality(string qualityMode)
    {
        return qualityMode.ToLowerInvariant() switch
        {
            "keep" => "0",
            "compact" => "5",
            _ => "2"
        };
    }

    private static string GetMp4Crf(string qualityMode)
    {
        return qualityMode.ToLowerInvariant() switch
        {
            "keep" => "18",
            "compact" => "28",
            _ => "23"
        };
    }

    private static void ConvertWithLibreOffice(string sourcePath, string outputPath, CancellationToken cancellationToken)
    {
        var sofficePath = TryFindExecutable("soffice.exe", Array.Empty<string>())
            ?? throw new InvalidOperationException("LibreOffice was not found on this machine.");

        var tempDirectory = Path.Combine(Path.GetTempPath(), "argus-convert-" + Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture));
        Directory.CreateDirectory(tempDirectory);

        try
        {
            RunProcess(sofficePath, $"--headless --convert-to pdf --outdir {Quote(tempDirectory)} {Quote(sourcePath)}", cancellationToken);
            var expected = Path.Combine(tempDirectory, Path.GetFileNameWithoutExtension(sourcePath) + ".pdf");
            if (!File.Exists(expected))
            {
                throw new InvalidOperationException("LibreOffice did not create the expected PDF output.");
            }

            File.Copy(expected, outputPath, overwrite: false);
        }
        finally
        {
            TryDeleteDirectory(tempDirectory);
        }
    }

    private static void ConvertWithImageMagick(string sourcePath, string outputPath, CancellationToken cancellationToken)
    {
        var magickPath = TryFindExecutable("magick.exe", Array.Empty<string>())
            ?? throw new InvalidOperationException("ImageMagick was not found on this machine.");

        RunProcess(magickPath, $"{Quote(sourcePath)} {Quote(outputPath)}", cancellationToken);
    }

    private static void RunProcess(string fileName, string arguments, CancellationToken cancellationToken)
    {
        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            }
        };

        process.Start();
        var standardOutput = process.StandardOutput.ReadToEnd();
        var standardError = process.StandardError.ReadToEnd();
        process.WaitForExit();
        cancellationToken.ThrowIfCancellationRequested();

        if (process.ExitCode != 0)
        {
            var message = string.IsNullOrWhiteSpace(standardError) ? standardOutput : standardError;
            throw new InvalidOperationException(string.IsNullOrWhiteSpace(message)
                ? $"Conversion process failed with exit code {process.ExitCode}."
                : message.Trim());
        }
    }

    private static ConversionRecord CreateSkippedRecord(string sourcePath, string outputPath, string targetFormat, string converter, string message)
    {
        return new ConversionRecord
        {
            SourcePath = sourcePath,
            OutputPath = outputPath,
            TargetFormat = targetFormat,
            Converter = converter,
            Status = ConversionStatus.Skipped,
            Message = message
        };
    }

    private static string? TryFindExecutable(string executableName, IEnumerable<string> explicitCandidates)
    {
        foreach (var candidate in explicitCandidates.Where(path => !string.IsNullOrWhiteSpace(path)))
        {
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        var pathValue = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        foreach (var directory in pathValue.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
        {
            try
            {
                var candidate = Path.Combine(directory.Trim(), executableName);
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }
            catch
            {
                // Ignore malformed PATH entries.
            }
        }

        return null;
    }

    private static void TryDeleteFile(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
            // Best-effort cleanup only.
        }
    }

    private static void TryDeleteDirectory(string path)
    {
        try
        {
            if (Directory.Exists(path))
            {
                Directory.Delete(path, recursive: true);
            }
        }
        catch
        {
            // Best-effort cleanup only.
        }
    }

    private static string AppendStatusMessage(string? existing, string addition)
    {
        if (string.IsNullOrWhiteSpace(existing))
        {
            return addition;
        }

        return existing + " " + addition;
    }

    private static string Quote(string value)
    {
        return "\"" + value.Replace("\"", "\\\"", StringComparison.Ordinal) + "\"";
    }
}



