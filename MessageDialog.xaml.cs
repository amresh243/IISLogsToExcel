using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace IISLogsToExcel;

public enum DialogResults { No = 0, Yes = 1 }
public enum DialogTypes { Error = 0, Warning = 1, Info = 2, Question = 3 }

/// <summary>
/// Interaction logic for MessageDialog.xaml
/// </summary>
public partial class MessageDialog : Window
{
    private Window? _owner;
    private DialogResults _result = DialogResults.No;
    private readonly Dictionary<DialogTypes, BitmapImage> _icons = new Dictionary<DialogTypes, BitmapImage>
    {
        { DialogTypes.Info, new BitmapImage(new Uri("pack://application:,,,/res/info.png")) },
        { DialogTypes.Warning, new BitmapImage(new Uri("pack://application:,,,/res/warning.png")) },
        { DialogTypes.Error, new BitmapImage(new Uri("pack://application:,,,/res/error.png")) },
        { DialogTypes.Question, new BitmapImage(new Uri("pack://application:,,,/res/question.png")) }
    };

    public MessageDialog(Window owner)
    {
        _owner = owner;
        InitializeComponent();
    }

    public DialogResults Show(string message, string title, DialogTypes type = DialogTypes.Info, Image? icon = null)
    {
        dlgImage.Source = (icon == null) ? _icons[type] : icon.Source;

        switch(type)
        {
            case DialogTypes.Error:
            case DialogTypes.Warning:
            case DialogTypes.Info:
                yesButton.Visibility = Visibility.Hidden;
                closeButton.Visibility = Visibility.Visible;
                noButton.Visibility = Visibility.Hidden;
                closeButton.Focus();
                break;
            case DialogTypes.Question:
                yesButton.Visibility = Visibility.Visible;
                closeButton.Visibility = Visibility.Hidden;
                noButton.Visibility = Visibility.Visible;
                noButton.Focus();
                break;
        }

        AppTitle.Text = title;
        Message.Text = message;
        this.Owner = _owner;
        this.WindowStartupLocation = WindowStartupLocation.CenterOwner;
        this.ShowDialog();

        return _result;
    }

    public void ApplyTheme(bool isDark = false)
    {
        var foreColor = (isDark) ? Brushes.White : Brushes.Black;
        var backColor = (isDark) ? Brushes.Black : Brushes.White;

        Message.Foreground = foreColor;
        Message.Background = backColor;
        this.Background = backColor;
    }

    private void Yes_Click(object sender, RoutedEventArgs e)
    {
        _result = DialogResults.Yes;
        this.Hide();
    }

    private void No_Click(object sender, RoutedEventArgs e)
    {
        _result = DialogResults.No;
        this.Hide();
    }
}
