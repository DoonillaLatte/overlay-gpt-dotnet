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
using overlay_gpt;
using overlay_gpt.Network;
using Microsoft.AspNetCore.SignalR;
using System.Text.Json;

namespace overlay_gpt 
{
    public partial class MainWindow : Window
    {
        private WebSocketServer _webSocketServer;
        private SocketIOConnection _socketIOConnection;

        public MainWindow()
        {
            InitializeComponent();
            
            // 웹소켓 서버 초기화
            _webSocketServer = new WebSocketServer(8080);
            
            // Socket.IO 클라이언트는 WebSocketServer에서 생성된 것을 사용
            _socketIOConnection = null;
            
            // 로그 윈도우 표시
            LogWindow.Instance.Show();
            LogWindow.Instance.Log("MainWindow Loaded");
            
            // 창 숨기기 설정을 Loaded 이벤트 안으로 이동
            this.Loaded += async (s, e) =>
            {
                try
                {
                    await _webSocketServer.StartAsync();
                    _socketIOConnection = _webSocketServer.GetSocketIOConnection();
                    _socketIOConnection.OnMessageReceived += (sender, message) =>
                    {
                        LogWindow.Instance.Log($"Flask 서버로부터 메시지 수신: {message}");
                    };

                    Console.WriteLine("Loaded");
                    var helper = new WindowInteropHelper(this);
                    HotkeyManager.RegisterHotKey(helper, ShowOverlay);
                    
                    // 핫키 등록 후에 창 숨기기
                    this.Hide();
                    this.ShowInTaskbar = false;
                }
                catch (Exception ex)
                {
                    LogWindow.Instance.Log($"서버 시작 중 오류 발생: {ex.Message}");
                }
            };

            // 창이 닫힐 때 웹소켓 서버와 클라이언트 종료
            this.Closing += async (s, e) =>
            {
                await _webSocketServer.StopAsync();
                await _socketIOConnection.DisconnectAsync();
            };
        }

        private async void ShowOverlay()
        {
            try
            {
                var element = AutomationElement.FocusedElement;
                var reader = ContextReaderFactory.CreateReader(element);
                LogWindow.Instance.Log($"Reader Type: {reader.GetType().Name}");
                var result = reader.GetSelectedTextWithStyle();
                string context = result.SelectedText;
                
                LogWindow.Instance.Log($"Selected Text: {context}");

                var displayTextMessage = new
                {
                    command = "display_text",
                    chat_id = 1,
                    current_program = new
                    {
                        file_name = "",
                        program_type = "",
                        context = context
                    },
                    target_program = new
                    {
                        file_name = "",
                        program_type = "",
                        context = ""
                    },
                    texts = new[]
                    {
                        new
                        {
                            type = "text_plain",
                            content = context
                        }
                    }
                };

                Console.WriteLine("==========================================");
                Console.WriteLine("Vue로 메시지 전송 중...");
                Console.WriteLine($"전송 시간: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                Console.WriteLine($"전송할 메시지: {JsonSerializer.Serialize(displayTextMessage, new JsonSerializerOptions { WriteIndented = true })}");
                await _webSocketServer.GetHubContext().Clients.All.SendAsync("ReceiveMessage", displayTextMessage);
                Console.WriteLine("Vue로 메시지 전송 완료");
                Console.WriteLine("==========================================");
            }
            catch (Exception ex)
            {
                LogWindow.Instance.Log($"오류 발생: {ex.Message}");
                Console.WriteLine($"오류 발생: {ex.Message}");
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