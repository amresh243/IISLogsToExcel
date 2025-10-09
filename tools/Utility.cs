// Author: Amresh Kumar (July 2025)

using Microsoft.Win32;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace IISLogsToExcel.tools;

public static class Utility
{
    private const string _xmlPatterns = @"[\u0009\u000A\u000D\u0020-\uD7FF\uE000-\uFFFD]";
    private const string _numberPatterns = @"^\d+$";

    /// <summary> Returns a valid number from the given string. </summary>
    public static int GetValidNumber(this string text) =>
        int.TryParse(text, out int number) ? number : 0;

    /// <summary> Checks if the given string is numeric. </summary>
    public static bool IsNumeric(this string input) =>
        !string.IsNullOrEmpty(input) && !input.Any(static c => !char.IsDigit(c));

    /// <summary> Checks if the given string is numeric (slower). </summary>
    public static bool IsNumeric2(this string input) =>
        !string.IsNullOrEmpty(input) && Regex.IsMatch(input, _numberPatterns);

    /// <summary> Removes invalid XML characters from the given text. </summary>
    public static string RemoveInvalidXmlChars(this string text) =>
        string.IsNullOrEmpty(text)
            ? text 
            : new string([.. text.Where(ch =>
                ch == 0x9 || ch == 0xA || ch == 0xD ||
                ch >= 0x20 && ch <= 0xD7FF ||
                ch >= 0xE000 && ch <= 0xFFFD ||
                ch >= 0x10000 && ch <= 0x10FFFF)]);

    /// <summary> Removes invalid XML characters from the given text (slower) </summary>
    public static string RemoveInvalidXmlChars2(this string text) =>
        string.IsNullOrEmpty(text)
            ? text
            // Use regex to match and rebuild the string
            : string.Concat(Regex.Matches(text, _xmlPatterns).Cast<Match>().Select(m => m.Value));

    /// <summary> Returns all log files under the given folder path. </summary>
    public static string[] GetLogFiles(string folderPath, string extension = "*.log") =>
        string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath)
            ? []
            : Directory.GetFiles(folderPath, extension, SearchOption.AllDirectories);

    /// <summary> Checks if the system is in dark mode. </summary>
    public static bool IsSystemInDarkMode()
    {
        using RegistryKey? key = Registry.CurrentUser.OpenSubKey(Constants.ThemeKey);
        if (key != null)
        {
            object? registryValueObject = key.GetValue(Constants.ThemeValue);
            if (registryValueObject != null)
            {
                int registryValue = (int)registryValueObject;
                return registryValue == 0; // 0 = Dark Mode, 1 = Light Mode
            }
        }

        // Default to light mode if Key not found
        return false;
    }

    /// <summary> Returns the formatted size string for the given size in bytes. </summary>
    public static string GetFormattedSize(long size)
    {
        if (size < 1024)
            return $"{size} Bytes";
        else if (size < 1048576)
            return $"{size / 1024.0:F2} KB";
        else if (size < 1073741824)
            return $"{size / 1048576.0:F2} MB";
        else
            return $"{size / 1073741824.0:F2} GB";
    }

    /// <summary> Returns the file name without the root folder path. </summary>
    public static string GetFileNameWithoutRoot(string file, string root) =>
        !root.EndsWith('\\') ? file.Replace(root + "\\", string.Empty) : file.Replace(root, string.Empty);

    /// <summary> Create a vertical linear gradient brush from startColor to endColor </summary>
    public static LinearGradientBrush GetGradientBrush(Color startColor, Color endColor, double opacity = 0)
    {
        LinearGradientBrush brush = new()
        {
            StartPoint = new Point(0.5, 0),
            EndPoint = new Point(0.5, 1)
        };

        brush.GradientStops.Add(new GradientStop(startColor, 0.0));
        brush.GradientStops.Add(new GradientStop(endColor, 1.0));
        if (opacity > 0)
            brush.Opacity = opacity;

        return brush;
    }

    public static void SetCheckBoxStyle(CheckBox checkBox, Brush brush)
    {
        if (checkBox == null || checkBox.IsChecked == null)
            return;

        if (checkBox.IsChecked == false)
            brush = GetStyle("ControlDisabled");

        checkBox.ApplyTemplate();
        var foregroundPanel = checkBox.Template.FindName("ForegroundPanel", checkBox) as Border;
        if (foregroundPanel != null)
            foregroundPanel.Background = brush;

    }

    public static Brush GetStyle(string key)
    {
        Window wnd = Application.Current.MainWindow!;
        LinearGradientBrush brush = (LinearGradientBrush)wnd.FindResource(key);
        return brush;
    }
}
