// Author: Amresh Kumar (August 2025)

using IISLogsToExcel.tools;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace IISLogsToExcel;

public enum DialogResults { No = 0, Yes = 1 }
public enum DialogTypes { Error = 0, Warning = 1, Info = 2, Question = 3 }

/// <summary> Interaction logic for MessageDialog.xaml </summary>
public partial class MessageDialog : Window
{
    private readonly Window? _owner;
    private DialogResults _result = DialogResults.No;

    private readonly Dictionary<DialogTypes, BitmapImage> _icons = new()
    {
        { DialogTypes.Info, new BitmapImage(new Uri(Icons.Info)) },
        { DialogTypes.Warning, new BitmapImage(new Uri(Icons.Warning)) },
        { DialogTypes.Error, new BitmapImage(new Uri(Icons.Error)) },
        { DialogTypes.Question, new BitmapImage(new Uri(Icons.Question)) }
    };

    private readonly Dictionary<DialogTypes, Brush>? _titleBarColors;

    /// <summary> Constructor </summary>
    public MessageDialog(Window owner)
    {
        InitializeComponent();

        _owner = owner;
        this.WindowStartupLocation = WindowStartupLocation.CenterOwner;
        _titleBarColors = new Dictionary<DialogTypes, Brush>
        {
            { DialogTypes.Info, TitleBar.Background },
            { DialogTypes.Warning, Utility.GetGradientBrush(Colors.LightGoldenrodYellow, Colors.Goldenrod) },
            { DialogTypes.Error, Utility.GetGradientBrush(Colors.Gold, Colors.Crimson) },
            { DialogTypes.Question, TitleBar.Background }
        };
    }

    /// <summary> Show the dialog with specified message, title, type and icon </summary>
    public DialogResults Show(string message, string title, DialogTypes type = DialogTypes.Info, Image? icon = null)
    {
        dlgImage.Source = (icon == null) ? _icons[type] : icon.Source;
        TitleBar.Background = _titleBarColors?[type] ?? TitleBar.Background;

        switch (type)
        {
            case DialogTypes.Error:
            case DialogTypes.Warning:
            case DialogTypes.Info:
                yesButton.Visibility = noButton.Visibility = Visibility.Hidden;
                closeButton.Visibility = Visibility.Visible;
                closeButton.Focus();
                break;
            case DialogTypes.Question:
                yesButton.Visibility = noButton.Visibility = Visibility.Visible;
                closeButton.Visibility = Visibility.Hidden;
                noButton.Focus();
                break;
        }

        AppTitle.Text = title;
        Message.Text = message;
        this.Owner = _owner;
        this.ShowDialog();

        return _result;
    }

    /// <summary> Apply theme colors </summary>
    public void ApplyTheme(Brush backColor, Brush foreColor)
    {
        Message.Foreground = foreColor;
        Message.Background = backColor;
        this.Background = backColor;
    }

    /// <summary> Yes button click event handler </summary>
    private void Yes_Click(object sender, RoutedEventArgs e)
    {
        _result = DialogResults.Yes;
        this.Hide();
    }

    /// <summary> No, Close and X button click event handler </summary>
    private void No_Click(object sender, RoutedEventArgs e)
    {
        _result = DialogResults.No;
        this.Hide();
    }

    /// <summary> Allow window drag on title bar mouse down </summary>
    private void Window_MouseDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton == MouseButton.Left)
            this.DragMove();
    }
}
