using System.ComponentModel;
using System.Windows.Media;

namespace IISLogsToExcel.tools
{
    internal class ColorItem
    {
        public ColorItem(Brush color, string name)
        {
            _ColorBrush = color;
            _Name = name;
        }

        public ColorItem() : this(Utility.GetGradientBrush(Colors.LightSkyBlue, Colors.DeepSkyBlue), "Default") { }

        /// <summary> Property to access ID </summary>
        public Brush ColorBrush
        {
            get => _ColorBrush;
            set
            {
                if (_ColorBrush != value)
                {
                    _ColorBrush = value;
                    RaisePropertyChanged(nameof(_ColorBrush));
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

        protected void RaisePropertyChanged(string propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        public event PropertyChangedEventHandler? PropertyChanged;

        private Brush _ColorBrush = Utility.GetGradientBrush(Colors.LightSkyBlue, Colors.DeepSkyBlue);
        private string _Name = string.Empty;
    }
}
