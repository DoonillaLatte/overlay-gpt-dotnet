using System.Windows;

namespace overlay_gpt;

public partial class LogWindow : Window
{
    private static LogWindow? _instance;

    public static LogWindow Instance
    {
        get
        {
            if (_instance == null || !_instance.IsLoaded)
            {
                _instance = new LogWindow();
            }
            return _instance;
        }
    }

    private LogWindow()
    {
        InitializeComponent();
    }

    public void Log(string message)
    {
        Dispatcher.Invoke(() =>
        {
            LogTextBox.AppendText($"{message}\n");
            LogTextBox.ScrollToEnd();
        });
    }
} 