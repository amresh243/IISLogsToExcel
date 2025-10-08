// Author: Amresh Kumar (July 2025)

using System.ComponentModel;
using System.Windows.Media;

namespace IISLogsToExcel.tools;

public class LogFileItem : INotifyPropertyChanged
{
    public LogFileItem(string id, string name, string fullPath, string toolTip, Brush color, Brush idColor)
    {
        _ID = id;
        _Name = name;
        FullPath = fullPath;
        _ToolTip = toolTip;
        _Color = color;
        _IdColor = idColor;
    }

    public LogFileItem() : this(string.Empty, string.Empty, string.Empty, string.Empty, Brushes.Black, Brushes.Black) {}

    /// <summary> Property to access ID </summary>
    public string ID
    {
        get => $"{_ID}. ";
        set
        {
            if (_ID != value)
            {
                _ID = value;
                RaisePropertyChanged(nameof(ID));
            }
        }
    }

    /// <summary> Returns indigo color if pattern is not standard </summary>
    public Brush IdColor
    {
        get => _IdColor;
        set
        {
            if (_IdColor != value)
            {
                _IdColor = value;
                RaisePropertyChanged(nameof(IdColor));
            }
        }
    }

    /// <summary> Property to access file name </summary>
    public string Name
    {
        get => _Name;
        set
        {
            if (_Name != value)
            {
                _Name = value;
                RaisePropertyChanged(nameof(Name));
            }
        }
    }

    /// <summary> Property to access full path </summary>
    public string FullPath { get; set; }

    /// <summary> Property to access tooltip </summary>
    public string ToolTip
    {         
        get => _ToolTip;
        set
        {
            if (_ToolTip != value)
            {
                _ToolTip = value;
                RaisePropertyChanged(nameof(ToolTip));
            }
        }
    }

    /// <summary> Returns indigo color if pattern is not standard </summary>
    public Brush Color
    {
        get => _Color;
        set
        {
            if (_Color != value)
            {
                _Color = value;
                RaisePropertyChanged(nameof(Color));
            }
        }
    }

    protected void RaisePropertyChanged(string propertyName) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

    public event PropertyChangedEventHandler? PropertyChanged;

    private string _ID = string.Empty;
    private string _Name = string.Empty;
    private string _ToolTip = string.Empty;
    private Brush _Color = Brushes.Black;
    private Brush _IdColor = Brushes.Black;
}
