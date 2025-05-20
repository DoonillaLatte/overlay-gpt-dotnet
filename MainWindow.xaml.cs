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
using System.Windows.Interop;
using System.Windows.Automation;

namespace overlay_gpt 
{
    public partial class MainWindow : Window
    {
        private OverlayWindow? overlayWindow;
        
        public MainWindow()
        {
            InitializeComponent();
            
            
            // 로그 윈도우 표시
            LogWindow.Instance.Show();
            LogWindow.Instance.Log("MainWindow Loaded");
            
            // 창 숨기기 설정을 Loaded 이벤트 안으로 이동
            this.Loaded += (s, e) =>
            {
                Console.WriteLine("Loaded");
                var helper = new WindowInteropHelper(this);
                HotkeyManager.RegisterHotKey(helper, ShowOverlay);
                
                // 핫키 등록 후에 창 숨기기
                this.Hide();
                this.ShowInTaskbar = false;
            };
        }

        private void ShowOverlay()
        {
            var element = AutomationElement.FocusedElement;
            var reader = ContextReaderFactory.CreateReader(element);
            LogWindow.Instance.Log($"Reader Type: {reader.GetType().Name}");
            var (selectedText, _) = reader.GetSelectedTextWithStyle();
            string context = selectedText;
            
            if (overlayWindow == null || !overlayWindow.IsLoaded)
            {
                overlayWindow = new OverlayWindow(context);
                overlayWindow.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            }
            else
            {
                // 이미 보이는 상태면 숨기고, 숨겨진 상태면 보이기
                if (overlayWindow.IsVisible)
                {
                    overlayWindow.Hide();
                }
                else
                {
                    overlayWindow.Show();
                    overlayWindow.Activate();
                }
            }
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            // 창을 최소화해도 핫키가 동작하도록 설정
            WindowState = WindowState.Minimized;
        }
    }
}