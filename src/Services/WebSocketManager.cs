using System;
using System.Threading.Tasks;
using System.Text.Json;

namespace overlay_gpt.Services
{
    public class WebSocketManager
    {
        private static WebSocketManager? _instance;
        private readonly WebSocketService _webSocketService;
        private readonly string _serverUrl;
        private bool _isInitialized;

        private WebSocketManager(string serverUrl)
        {
            _serverUrl = serverUrl;
            _webSocketService = new WebSocketService();
            _isInitialized = false;
        }

        public static WebSocketManager Instance
        {
            get
            {
                if (_instance == null)
                {
                    throw new InvalidOperationException("WebSocketManager가 초기화되지 않았습니다. Initialize()를 먼저 호출하세요.");
                }
                return _instance;
            }
        }

        public static void Initialize(string serverUrl)
        {
            if (_instance != null)
            {
                throw new InvalidOperationException("WebSocketManager가 이미 초기화되었습니다.");
            }
            _instance = new WebSocketManager(serverUrl);
        }

        public async Task StartAsync()
        {
            if (_isInitialized)
            {
                return;
            }

            try
            {
                // 기본 명령어 핸들러 등록
                RegisterDefaultHandlers();

                // 웹소켓 서버 연결
                await _webSocketService.StartAsync(_serverUrl);
                _isInitialized = true;
                Console.WriteLine("웹소켓 서비스가 성공적으로 시작되었습니다.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"웹소켓 서비스 시작 중 오류 발생: {ex.Message}");
                throw;
            }
        }

        private void RegisterDefaultHandlers()
        {
            // 연결 상태 변경 핸들러
            _webSocketService.RegisterCommandHandler("connectionStatus", async (parameters) =>
            {
                var status = parameters.GetProperty("status").GetString();
                Console.WriteLine($"연결 상태 변경: {status}");
            });

            // 에러 핸들러
            _webSocketService.RegisterCommandHandler("error", async (parameters) =>
            {
                var errorCode = parameters.GetProperty("code").GetInt32();
                var errorMessage = parameters.GetProperty("message").GetString();
                Console.WriteLine($"에러 발생: {errorCode} - {errorMessage}");
            });
        }

        public void RegisterCommandHandler(string commandName, Func<JsonElement, Task> handler)
        {
            _webSocketService.RegisterCommandHandler(commandName, handler);
        }

        public async Task SendCommandAsync(string commandName, object parameters)
        {
            if (!_isInitialized)
            {
                throw new InvalidOperationException("웹소켓 서비스가 초기화되지 않았습니다.");
            }

            await _webSocketService.SendCommandAsync(commandName, parameters);
        }

        public async Task StopAsync()
        {
            if (!_isInitialized)
            {
                return;
            }

            await _webSocketService.StopAsync();
            _isInitialized = false;
        }
    }
} 