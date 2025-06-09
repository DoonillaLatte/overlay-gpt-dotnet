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
using overlay_gpt.Network.Models.Vue;
using overlay_gpt.Services;
using Microsoft.AspNetCore.SignalR;
using System.Text.Json;

namespace overlay_gpt 
{
    public partial class MainWindow : Window
    {
        private WithVueAsHost _vueServer;
        private WithFlaskAsClient _flaskClient;
        private TextProcessingService _textProcessingService;
        private DateTime _lastFetchTime = DateTime.MinValue;
        private const int FetchDebounceMs = 2000; // 2초

        public static MainWindow Instance { get; private set; }
        public WithVueAsHost VueServer => _vueServer;
        public WithFlaskAsClient FlaskClient => _flaskClient;

        public MainWindow()
        {
            InitializeComponent();
            Instance = this;
            
            // Flask 클라이언트 초기화
            _flaskClient = new WithFlaskAsClient();
            
            // Vue 서버 초기화 (Flask 클라이언트 전달)
            _vueServer = new WithVueAsHost(8080, _flaskClient);
            
            // 텍스트 처리 서비스 초기화
            _textProcessingService = new TextProcessingService();
            
            // 로그 윈도우 표시
            LogWindow.Instance.Show();
            LogWindow.Instance.Log("MainWindow Loaded");
            
            // 창 숨기기 설정을 Loaded 이벤트 안으로 이동
            this.Loaded += async (s, e) =>
            {
                try
                {
                    await _vueServer.StartAsync();
                    await _flaskClient.ConnectAsync();
                    
                    // LogWindow에 서버 인스턴스 전달
                    LogWindow.Instance.SetServers(_flaskClient, _vueServer);
                    
                    _flaskClient.On("message", (response) =>
                    {
                        LogWindow.Instance.Log($"Flask 서버로부터 메시지 수신: {response.GetValue<string>()}");
                    });

                    Console.WriteLine("Loaded");
                    var helper = new WindowInteropHelper(this);
                    HotkeyManager.RegisterHotKey(helper, FetchContext);
                    
                    // 핫키 등록 후에 창 숨기기
                    this.Hide();
                    this.ShowInTaskbar = false;
                }
                catch (Exception ex)
                {
                    LogWindow.Instance.Log($"서버 시작 중 오류 발생: {ex.Message}");
                }
            };

            // 창이 닫힐 때 서버와 클라이언트 종료
            this.Closing += async (s, e) =>
            {
                await _vueServer.StopAsync();
                await _flaskClient.DisconnectAsync();
            };
        }

        private async void FetchContext()
        {
            try
            {
                // Debounce 로직: 마지막 호출로부터 2초 이내면 무시
                var now = DateTime.Now;
                var timeSinceLastFetch = (now - _lastFetchTime).TotalMilliseconds;
                
                if (timeSinceLastFetch < FetchDebounceMs)
                {
                    LogWindow.Instance.Log($"중복 요청 방지: 마지막 요청으로부터 {timeSinceLastFetch:F0}ms 경과 (최소 {FetchDebounceMs}ms 필요)");
                    return;
                }
                
                _lastFetchTime = now;
                LogWindow.Instance.Log($"[{now:yyyy-MM-dd HH:mm:ss.fff}] 컨텍스트 가져오기 시작");
                
                var element = AutomationElement.FocusedElement;
                var reader = ContextReaderFactory.CreateReader(element);
                LogWindow.Instance.Log($"Reader Type: {reader.GetType().Name}");
                var result = reader.GetSelectedTextWithStyle();
                string context = result.SelectedText;
                
                LogWindow.Instance.Log($"Selected Text: {context}");

                // 파일 정보 가져오기
                var fileInfo = reader.GetFileInfo();
                LogWindow.Instance.Log($"File Info - ID: {fileInfo.FileId}, Volume: {fileInfo.VolumeId}, Type: {fileInfo.FileType}, Name: {fileInfo.FileName}, Path: {fileInfo.FilePath}");
                
                var programInfo = new Network.Models.Common.ProgramInfo
                {
                    Context = context,
                    FileId = fileInfo.FileId,
                    VolumeId = fileInfo.VolumeId,
                    FileType = fileInfo.FileType,
                    FileName = fileInfo.FileName,
                    FilePath = fileInfo.FilePath,
                    Position = result.LineNumber
                };
                
                LogWindow.Instance.Log($"Program Info - ID: {programInfo.FileId}, Volume: {programInfo.VolumeId}, Type: {programInfo.FileType}, Name: {programInfo.FileName}, Path: {programInfo.FilePath}");
                
                var displayTextMessage = _textProcessingService.ProcessSelectedText(context, programInfo);

                Console.WriteLine("==========================================");
                Console.WriteLine("Vue로 메시지 전송 중...");
                Console.WriteLine($"전송 시간: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                Console.WriteLine($"전송할 메시지: {_textProcessingService.SerializeMessage(displayTextMessage)}");
                await _vueServer.SendMessageToAll(displayTextMessage);
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