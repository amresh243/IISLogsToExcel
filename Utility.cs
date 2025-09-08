// Author: Amresh Kumar (July 2025)

using Microsoft.Win32;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Media;

namespace IISLogsToExcel;

public static class Utility
{
    private const string _xmlPatterns = @"[\u0009\u000A\u000D\u0020-\uD7FF\uE000-\uFFFD]";
    private const string _numberPatterns = @"^\d+$";

    /// <summary> Returns a valid number from the given string. </summary>
    public static int GetValidNumber(this string text)
    {
        if (int.TryParse(text, out int number))
            return number;

        return 0;
    }

    /// <summary> Checks if the given string is numeric. </summary>
    public static bool IsNumeric(this string input)
    {
        if (string.IsNullOrEmpty(input) || input.Any(c => !char.IsDigit(c)))
            return false;

        return true;
    }

    /// <summary> Checks if the given string is numeric (slower). </summary>
    public static bool IsNumeric2(this string input)
    {
        if (string.IsNullOrEmpty(input))
            return false;

        return Regex.IsMatch(input, _numberPatterns);
    }

    /// <summary> Removes invalid XML characters from the given text. </summary>
    /// <param name="text">Input text</param>
    /// <returns>Cleaned text</returns>
    public static string RemoveInvalidXmlChars(this string text)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        return new string([.. text.Where(ch =>
            (ch == 0x9 || ch == 0xA || ch == 0xD) ||
            (ch >= 0x20 && ch <= 0xD7FF) ||
            (ch >= 0xE000 && ch <= 0xFFFD) ||
            (ch >= 0x10000 && ch <= 0x10FFFF))]);
    }

    /// <summary> Removes invalid XML characters from the given text (slower) </summary>
    /// <param name="text">Input text</param>
    /// <returns>Cleaned text</returns>
    public static string RemoveInvalidXmlChars2(this string text)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        // Use regex to match and rebuild the string
        MatchCollection matches = Regex.Matches(text, _xmlPatterns);
        return string.Concat(matches.Cast<Match>().Select(m => m.Value));

    }

    /// <summary> Returns all log files under the given folder path. </summary>
    /// <param name="folderPath">Log folder path.</param>
    /// <returns>Array of list file paths.</returns>
    public static string[] GetLogFiles(string folderPath, string extension = "*.log")
    {
        if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
            return [];

        return Directory.GetFiles(folderPath, extension, SearchOption.AllDirectories);
    }

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

    public static string GetFileNameWithoutRoot(string file, string root) =>
        !root.EndsWith('\\') ? file.Replace(root + "\\", string.Empty) : file.Replace(root, string.Empty);

    /// <summary> Create a vertical linear gradient brush from startColor to endColor </summary>
    public static LinearGradientBrush GetLinearGradientBrush(Color startColor, Color endColor, double opacity = 0)
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
}
