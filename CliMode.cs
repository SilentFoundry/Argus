using System;
using System.Collections.Generic;
using System.Linq;

namespace FileOrganizer;

internal static class CliMode
{
    public static int Run(string[] args)
    {
        try
        {
            var parsed = Parse(args);
            if (parsed.ShowHelp)
            {
                PrintHelp();
                return 0;
            }

            if (parsed.UndoLastRun)
            {
                var undoResult = UndoService.UndoLastRun();
                Console.WriteLine(undoResult.DetailedReport);
                return 0;
            }

            var result = OrganizerCore.Execute(parsed.Options);
            Console.WriteLine(result.DetailedReport);
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.Message);
            return 1;
        }
    }

    private static ParsedArguments Parse(string[] args)
    {
        var options = new OrganizerOptions();
        var showHelp = false;
        var undoLastRun = false;

        for (var index = 0; index < args.Length; index++)
        {
            var argument = args[index].Trim();
            if (string.IsNullOrWhiteSpace(argument))
            {
                continue;
            }

            if (!argument.StartsWith("-", StringComparison.Ordinal) &&
                !argument.StartsWith("/", StringComparison.Ordinal))
            {
                if (string.IsNullOrWhiteSpace(options.SourcePath))
                {
                    options = options with { SourcePath = argument };
                    continue;
                }

                throw new InvalidOperationException($"Unexpected argument: {argument}");
            }

            switch (argument.ToLowerInvariant())
            {
                case "--help":
                case "-help":
                case "/help":
                case "-h":
                case "/?":
                    showHelp = true;
                    break;
                case "--undo-last-run":
                    undoLastRun = true;
                    break;
                case "--source":
                case "-source":
                case "-sourcepath":
                case "/source":
                case "/sourcepath":
                case "-s":
                    options = options with { SourcePath = GetValue(args, ref index, argument) };
                    break;
                case "--destination":
                case "-destination":
                case "-destinationpath":
                case "/destination":
                case "/destinationpath":
                    options = options with { DestinationPath = GetValue(args, ref index, argument) };
                    break;
                case "--mode":
                case "-mode":
                case "/mode":
                    options = options with { Mode = ParseOperationMode(GetValue(args, ref index, argument)) };
                    break;
                case "--organization-mode":
                    options = options with { OrganizationMode = ParseOrganizationMode(GetValue(args, ref index, argument)) };
                    break;
                case "--duplicate-handling":
                    options = options with { DuplicateHandling = ParseDuplicateHandling(GetValue(args, ref index, argument)) };
                    break;
                case "--metadata-rename-mode":
                case "--metadata-rename":
                case "--pdf-rename-mode":
                case "--pdf-rename":
                    options = options with { PdfRenameMode = ParsePdfRenameMode(GetValue(args, ref index, argument)) };
                    break;
                case "--rules":
                    options = options with { RulesPath = GetValue(args, ref index, argument) };
                    break;
                case "--tag-manifest-path":
                    options = options with { TagManifestPath = GetValue(args, ref index, argument) };
                    break;
                case "--include-extensions":
                case "--extensions":
                    options = options with { IncludeExtensions = ParseExtensions(GetValue(args, ref index, argument)) };
                    break;
                case "--top-level-only":
                    options = options with { TopLevelOnly = true };
                    break;
                case "--use-extension-subfolders":
                    options = options with { UseExtensionSubfolders = true };
                    break;
                case "--write-tag-manifest":
                    options = options with { WriteTagManifest = true };
                    break;
                case "--what-if":
                case "--whatif":
                    options = options with { WhatIf = true };
                    break;
                case "--clean-empty-folders":
                    options = options with { CleanEmptyFolders = true };
                    break;
                default:
                    throw new InvalidOperationException($"Unknown argument: {argument}");
            }
        }

        return new ParsedArguments(options, showHelp, undoLastRun);
    }

    private static string GetValue(string[] args, ref int index, string argument)
    {
        if (index + 1 >= args.Length)
        {
            throw new InvalidOperationException($"Missing value for {argument}");
        }

        index++;
        return args[index];
    }

    private static OperationMode ParseOperationMode(string value)
    {
        return value.ToLowerInvariant() switch
        {
            "move" => OperationMode.Move,
            "copy" => OperationMode.Copy,
            _ => throw new InvalidOperationException($"Invalid mode '{value}'. Use Move or Copy.")
        };
    }

    private static OrganizationMode ParseOrganizationMode(string value)
    {
        return value.ToLowerInvariant() switch
        {
            "bytype" or "type" => OrganizationMode.ByType,
            "byextension" or "extension" => OrganizationMode.ByExtension,
            "byyearmonth" or "yearmonth" => OrganizationMode.ByYearMonth,
            "bydatemodified" or "datemodified" => OrganizationMode.ByDateModified,
            "bysize" or "size" => OrganizationMode.BySize,
            "bykeyword" or "keyword" => OrganizationMode.ByKeyword,
            "categoryyearmonth" or "mixed" => OrganizationMode.CategoryYearMonth,
            _ => throw new InvalidOperationException($"Invalid organization mode '{value}'.")
        };
    }

    private static DuplicateHandlingMode ParseDuplicateHandling(string value)
    {
        return value.ToLowerInvariant() switch
        {
            "rename" => DuplicateHandlingMode.Rename,
            "skip" => DuplicateHandlingMode.Skip,
            "overwrite" => DuplicateHandlingMode.Overwrite,
            "keepnewest" => DuplicateHandlingMode.KeepNewest,
            "comparebyhash" or "hash" => DuplicateHandlingMode.CompareByHash,
            _ => throw new InvalidOperationException($"Invalid duplicate handling mode '{value}'.")
        };
    }

    private static PdfRenameMode ParsePdfRenameMode(string value)
    {
        return value.ToLowerInvariant() switch
        {
            "disabled" or "off" => PdfRenameMode.Disabled,
            "title" or "titleonly" => PdfRenameMode.TitleOnly,
            "content" or "contentheading" or "heading" => PdfRenameMode.ContentHeadingOnly,
            "titlethencontent" or "auto" => PdfRenameMode.TitleThenContent,
            _ => throw new InvalidOperationException($"Invalid metadata rename mode '{value}'.")
        };
    }

    private static void PrintHelp()
    {
        Console.WriteLine("Organize-Files");
        Console.WriteLine();
        Console.WriteLine("Launch without arguments to open the UI for the current build.");
        Console.WriteLine();
        Console.WriteLine("Web-shell build options:");
        Console.WriteLine("  --web");
        Console.WriteLine("  --no-browser");
        Console.WriteLine("  --port <number>");
        Console.WriteLine();
        Console.WriteLine("CLI options:");
        Console.WriteLine("  --source <folder>");
        Console.WriteLine("  --destination <folder>");
        Console.WriteLine("  --mode <move|copy>");
        Console.WriteLine("  --organization-mode <type|extension|yearmonth|datemodified|size|keyword|mixed>");
        Console.WriteLine("  --duplicate-handling <rename|skip|overwrite|keepnewest|hash>");
        Console.WriteLine("  --metadata-rename-mode <disabled|title|content|titlethencontent>");
        Console.WriteLine("    Alias: --pdf-rename-mode");
        Console.WriteLine("  --rules <file>");
        Console.WriteLine("  --tag-manifest-path <file>");
        Console.WriteLine("  --include-extensions <.pdf,.jpg,...>");
        Console.WriteLine("  --top-level-only");
        Console.WriteLine("  --use-extension-subfolders");
        Console.WriteLine("  --write-tag-manifest");
        Console.WriteLine("  --clean-empty-folders");
        Console.WriteLine("  --what-if");
        Console.WriteLine("  --undo-last-run");
        Console.WriteLine("  --help");
    }

    private static IReadOnlyList<string> ParseExtensions(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return Array.Empty<string>();
        }

        return value
            .Split(new[] { ',', ';', '\r', '\n', '\t', ' ' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(part => part.Trim())
            .Where(part => !string.IsNullOrWhiteSpace(part))
            .Select(part => part.StartsWith(".", StringComparison.Ordinal) ? part.ToLowerInvariant() : $".{part.ToLowerInvariant()}")
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private readonly record struct ParsedArguments(OrganizerOptions Options, bool ShowHelp, bool UndoLastRun);
}
