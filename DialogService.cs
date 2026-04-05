using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace FileOrganizer;

internal static class DialogService
{
    public static string? PickFolder(string? initialPath, string description)
    {
        return RunSta(() =>
        {
            using var dialog = new FolderBrowserDialog
            {
                Description = description,
                ShowNewFolderButton = true
            };

            if (!string.IsNullOrWhiteSpace(initialPath) && Directory.Exists(initialPath))
            {
                dialog.SelectedPath = initialPath;
            }

            return dialog.ShowDialog() == DialogResult.OK ? dialog.SelectedPath : null;
        });
    }

    public static string? PickSaveFile(string? initialPath, string title, string filter, string defaultFileName)
    {
        return RunSta(() =>
        {
            using var dialog = new SaveFileDialog
            {
                Title = title,
                Filter = filter,
                FileName = GetFileNameOrDefault(initialPath, defaultFileName)
            };

            var initialDirectory = ResolveInitialDirectory(initialPath);
            if (!string.IsNullOrWhiteSpace(initialDirectory))
            {
                dialog.InitialDirectory = initialDirectory;
            }

            return dialog.ShowDialog() == DialogResult.OK ? dialog.FileName : null;
        });
    }

    public static string? PickOpenFile(string? initialPath, string title, string filter)
    {
        return RunSta(() =>
        {
            using var dialog = new OpenFileDialog
            {
                Title = title,
                Filter = filter,
                CheckFileExists = false
            };

            if (!string.IsNullOrWhiteSpace(initialPath))
            {
                var initialDirectory = ResolveInitialDirectory(initialPath);
                if (!string.IsNullOrWhiteSpace(initialDirectory))
                {
                    dialog.InitialDirectory = initialDirectory;
                }

                if (File.Exists(initialPath))
                {
                    dialog.FileName = initialPath;
                }
            }

            return dialog.ShowDialog() == DialogResult.OK ? dialog.FileName : null;
        });
    }

    private static string GetFileNameOrDefault(string? path, string defaultFileName)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return defaultFileName;
        }

        var fileName = Path.GetFileName(path);
        return string.IsNullOrWhiteSpace(fileName) ? defaultFileName : fileName;
    }

    private static string? ResolveInitialDirectory(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return null;
        }

        try
        {
            if (Directory.Exists(path))
            {
                return path;
            }

            var directory = Path.GetDirectoryName(path);
            return !string.IsNullOrWhiteSpace(directory) && Directory.Exists(directory)
                ? directory
                : null;
        }
        catch
        {
            return null;
        }
    }

    private static T RunSta<T>(Func<T> action)
    {
        T? result = default;
        Exception? exception = null;

        var thread = new Thread(() =>
        {
            try
            {
                result = action();
            }
            catch (Exception ex)
            {
                exception = ex;
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();

        if (exception is not null)
        {
            throw exception;
        }

        return result!;
    }
}
