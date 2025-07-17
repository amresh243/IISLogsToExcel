// Author: Amresh Kumar (July 2025)

using System.IO;

namespace IISLogsToExcel
{
    public class IniFile
    {
        private readonly Dictionary<string, Dictionary<string, string>> _data = new();
        private readonly string _filePath;

        public IniFile(string filePath)
        {
            _filePath = filePath;
            if (File.Exists(filePath))
                Load(filePath);
        }

        private void Load(string path)
        {
            try
            {
                string? currentSection = null;

                foreach (var line in File.ReadAllLines(path))
                {
                    var trimmed = line.Trim();

                    if (string.IsNullOrWhiteSpace(trimmed) || trimmed.StartsWith(";"))
                        continue;

                    if (trimmed.StartsWith("[") && trimmed.EndsWith("]"))
                    {
                        currentSection = trimmed[1..^1].Trim();
                        if (!_data.ContainsKey(currentSection))
                            _data[currentSection] = new Dictionary<string, string>();
                    }
                    else if (trimmed.Contains('=') && currentSection != null)
                    {
                        var parts = trimmed.Split('=', 2);
                        var key = parts[0].Trim();
                        var value = parts[1].Trim();
                        _data[currentSection][key] = value;
                    }
                }
            }
            catch
            {
                // nothing to do here
            }
        }

        public string? GetValue(string section, string key)
        {
            return _data.TryGetValue(section, out var sectionData) && sectionData.TryGetValue(key, out var value)
                ? value
                : null;
        }

        public void SetValue(string section, string key, string value)
        {
            if (!_data.ContainsKey(section))
                _data[section] = new Dictionary<string, string>();

            _data[section][key] = value;
        }

        public bool Save()
        {
            try
            {
                using var writer = new StreamWriter(_filePath);
                foreach (var section in _data)
                {
                    writer.WriteLine($"[{section.Key}]");
                    foreach (var kvp in section.Value)
                    {
                        writer.WriteLine($"{kvp.Key}={kvp.Value}");
                    }
                    writer.WriteLine();
                }

                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}