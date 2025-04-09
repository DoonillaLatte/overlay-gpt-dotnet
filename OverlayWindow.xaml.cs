using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Automation;

namespace overlay_gpt;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class OverlayWindow : Window
{
    private string? _context;

    public OverlayWindow(string? context = null)
    {
        InitializeComponent();
        _context = context;
        
        // ESC 키 이벤트 처리
        this.PreviewKeyDown += (s, e) =>
        {
            if (e.Key == Key.Escape)
            {
                this.Hide();
            }
        };

        this.Loaded += (s, e) => {
            LogWindow.Instance.Log("OverlayWindow Loaded");
            if (!string.IsNullOrEmpty(_context))
            {
                inputTextBox.Text = _context;
            }
            else 
            {
                var element = AutomationElement.FocusedElement;
                var reader = ContextReaderFactory.CreateReader(element);
                var (selectedText, _) = reader.GetSelectedTextWithStyle();
                if (!string.IsNullOrEmpty(selectedText))
                {
                    inputTextBox.Text = selectedText;
                }
                else 
                {
                    inputTextBox.Text = "No Text";
                }
            }
            inputTextBox.Focus();
        };
    }
}