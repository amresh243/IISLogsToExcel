// Author: Amresh Kumar (July 2025)

using System.IO;
using System.Windows;

namespace IISLogsToExcel.Tools;

public enum DataType { STRING, INT, BOOL, DOUBLE };

public class IniFile
{
    private readonly Dictionary<string, Dictionary<string, string>> _data = [];
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

                if (string.IsNullOrWhiteSpace(trimmed) || trimmed.StartsWith(';'))
                    continue;

                if (trimmed.StartsWith('[') && trimmed.EndsWith(']'))
                {
                    currentSection = trimmed[1..^1].Trim();
                    if (!_data.ContainsKey(currentSection))
                        _data[currentSection] = [];
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
        catch { /* nothing to do here */ }
    }

    public string? GetValue(string section, string key) =>
        _data.TryGetValue(section, out var sectionData) && sectionData.TryGetValue(key, out var value) ? value : null;

    public object? GetTypedValue(string section, string key, DataType type)
    {
        var data = GetValue(section, key);
        if (data == null) 
            return null;

        try
        {
            return type switch
            {
                DataType.STRING => data,
                DataType.BOOL => bool.Parse(data),
                DataType.INT => int.Parse(data),
                DataType.DOUBLE => double.Parse(data),
                _ => data
            };
        } catch
        {
            return null;
        }
    }

    public void SetValue(string section, string key, string value)
    {
        if (!_data.ContainsKey(section))
            _data[section] = [];

        _data[section][key] = value;
    }

    private bool UpdateSettings()
    {
        using var writer = new StreamWriter(_filePath);
        foreach (var section in _data)
        {
            writer.WriteLine($"[{section.Key}]");
            foreach (var kvp in section.Value)
                writer.WriteLine($"{kvp.Key}={kvp.Value}");

            writer.WriteLine();
        }

        return true;
    }

    public bool Save(IISLogExporter? app = null)
    {
        try
        {
            return UpdateSettings();
        }
        catch
        {
            if(app == null)
                MessageBox.Show(Messages.SettingError, Captions.SettingError, MessageBoxButton.OK, MessageBoxImage.Warning);
            else
                app.MessageBox.Show(Messages.SettingError, Captions.SettingError, DialogTypes.Warning);

            return false;
        }
    }
}
