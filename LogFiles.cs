// Author: Amresh Kumar (July 2025)

using System.ComponentModel;
using System.Windows.Media;

namespace IISLogsToExcel
{
    public class LogFiles : INotifyPropertyChanged
    {
        /// <summary> Property to access ID </summary>
        public string ID
        {
            get { return $"{_ID}. "; }
            set
            {
                if (_ID != value)
                {
                    _ID = value;
                    RaisePropertyChanged(nameof(ID));
                }
            }
        }

        /// <summary> Property to access length of pattern </summary>
        public int Length => _Name.Length;

        /// <summary> Property to access existing pattern </summary>
        public string Name
        {
            get { return _Name; }
            set
            {
                if (_Name != value)
                {
                    _Name = value;
                    RaisePropertyChanged(nameof(Name));
                }
            }
        }

        /// <summary> Returns indigo color if pattern is not standard </summary>
        public Brush Color
        {
            get { return _Color; }
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
        private Brush _Color = Brushes.Black;
    }
}
