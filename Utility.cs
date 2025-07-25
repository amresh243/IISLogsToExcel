// Author: Amresh Kumar (July 2025)

using Microsoft.Win32;
using System.IO;

namespace IISLogsToExcel
{
    public static class Utility
    {
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
            if (string.IsNullOrEmpty(input))
                return false;

            var nonDigit = input.Where(c => !char.IsDigit(c)).ToList();
            if (nonDigit.Count > 0)
                return false;

            return true;
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
    }
}
