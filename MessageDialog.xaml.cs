// Author: Amresh Kumar (August 2025)

using IISLogsToExcel.tools;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace IISLogsToExcel;

public enum DialogResults { No = 0, Yes = 1 }
public enum DialogTypes { Error = 0, Warning = 1, Info = 2, Question = 3 }
public enum QuestionTypes { Info = 0, Warning = 1, Error = 2 }

/// <summary> Interaction logic for MessageDialog.xaml </summary>
public partial class MessageDialog : Window
{
    private DialogResults _result = DialogResults.No;
    private readonly Window? _owner;
    private readonly Dictionary<DialogTypes, LinearGradientBrush> _titleBarColors = new()
    {
        { DialogTypes.Info, Utility.GetGradientBrush(Colors.LightSkyBlue, Colors.DeepSkyBlue) },
        { DialogTypes.Warning, Utility.GetGradientBrush(Colors.LightGoldenrodYellow, Colors.Goldenrod) },
        { DialogTypes.Error, Utility.GetGradientBrush(Colors.Gold, Colors.Crimson) },
        { DialogTypes.Question, Utility.GetGradientBrush(Colors.LightSkyBlue, Colors.DeepSkyBlue) }
    };
    private readonly Dictionary<QuestionTypes, LinearGradientBrush> _questionTitleBarColor = new()
    {
        { QuestionTypes.Info, Utility.GetGradientBrush(Colors.LightSkyBlue, Colors.DeepSkyBlue) },
        { QuestionTypes.Warning, Utility.GetGradientBrush(Colors.LightGoldenrodYellow, Colors.Goldenrod) },
        { QuestionTypes.Error, Utility.GetGradientBrush(Colors.Gold, Colors.Crimson) }
    };
    private readonly Dictionary<DialogTypes, BitmapImage> _icons = new()
    {
        { DialogTypes.Info, new BitmapImage(new Uri(Icons.Info)) },
        { DialogTypes.Warning, new BitmapImage(new Uri(Icons.Warning)) },
        { DialogTypes.Error, new BitmapImage(new Uri(Icons.Error)) },
        { DialogTypes.Question, new BitmapImage(new Uri(Icons.Question)) }
    };

    /// <summary> Constructor </summary>
    public MessageDialog(Window owner)
    {
        InitializeComponent();

        _owner = owner;
        this.WindowStartupLocation = WindowStartupLocation.CenterOwner;
    }

    /// <summary> Show the dialog with specified message, title, type and icon </summary>
    public DialogResults Show(string message, string title, DialogTypes type = DialogTypes.Info, 
        Image? icon = null, QuestionTypes questionType = QuestionTypes.Info)
    {
        dlgImage.Source = (icon == null) ? _icons[type] : icon.Source;
        TitleBar.Background = _titleBarColors[type];

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
                if (questionType != QuestionTypes.Info)
                    TitleBar.Background = _questionTitleBarColor[questionType];

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
}
