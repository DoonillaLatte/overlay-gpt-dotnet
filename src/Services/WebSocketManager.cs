using System;
using System.Net;
using System.Net.WebSockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Text.Json;
using System.Collections.Generic;
using System.Windows;
using System.Linq;

namespace overlay_gpt.Services
{
    public class WebSocketManager
    {
        private static WebSocketManager? _instance;
        private HttpListener? _httpListener;
        private readonly string _url;
        private readonly WebSocketMessageHandler _messageHandler;
        private readonly LogWindow _logWindow;
        private CancellationTokenSource? _cancellationTokenSource;
        private bool _isRunning;
        private readonly List<WebSocket> _activeConnections = new List<WebSocket>();
        private readonly Dictionary<string, Action<JsonElement>> _messageHandlers = new Dictionary<string, Action<JsonElement>>();

        public event Action? OnConnected;

        public static WebSocketManager Instance
        {
            get
            {
                if (_instance == null)
                {
                    throw new InvalidOperationException("WebSocketManager가 초기화되지 않았습니다.");
                }
                return _instance;
            }
        }

        private WebSocketManager(string url, LogWindow logWindow)
        {
            if (!url.StartsWith("http://") && !url.StartsWith("https://"))
            {
                throw new ArgumentException("URL은 'http://' 또는 'https://'로 시작해야 합니다.");
            }

            if (!url.EndsWith("/"))
                url += "/";

            _url = url;
            _logWindow = logWindow;
            _messageHandler = new WebSocketMessageHandler(logWindow);
        }

        private void Log(string message)
        {
            MessageDispatcher.Instance.DispatchToUI(() => _logWindow.Log(message));
        }

        public static void Initialize(string url, LogWindow logWindow)
        {
            if (_instance != null)
            {
                throw new InvalidOperationException("WebSocketManager가 이미 초기화되었습니다.");
            }
            _instance = new WebSocketManager(url, logWindow);
        }

        public async Task StartAsync()
        {
            if (_isRunning)
            {
                return;
            }

            try
            {
                // 포트가 사용 중인지 확인
                var port = new Uri(_url).Port;
                var isPortInUse = System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties()
                    .GetActiveTcpListeners()
                    .Any(x => x.Port == port);

                if (isPortInUse)
                {
                    throw new InvalidOperationException($"포트 {port}가 이미 사용 중입니다.");
                }

                _httpListener = new HttpListener();
                _httpListener.Prefixes.Add(_url);
                _cancellationTokenSource = new CancellationTokenSource();

                _httpListener.Start();
                _isRunning = true;
                Log($"웹소켓 서버가 시작되었습니다: {_url}");

                while (!_cancellationTokenSource.Token.IsCancellationRequested)
                {
                    try
                    {
                        Log("클라이언트 연결 대기 중...");
                        var context = await _httpListener.GetContextAsync();
                        Log($"새로운 HTTP 요청 수신: {context.Request.Url}");
                        
                        if (context.Request.IsWebSocketRequest)
                        {
                            Log("웹소켓 업그레이드 요청 감지");
                            var webSocketContext = await context.AcceptWebSocketAsync(null);
                            Log("웹소켓 연결이 승인되었습니다.");
                            _ = HandleWebSocketConnectionAsync(webSocketContext.WebSocket);
                        }
                        else
                        {
                            Log("웹소켓이 아닌 HTTP 요청이 수신되었습니다.");
                            context.Response.StatusCode = 400;
                            context.Response.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"클라이언트 연결 처리 중 오류 발생: {ex.Message}");
                        Log($"스택 트레이스: {ex.StackTrace}");
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"웹소켓 서버 오류: {ex.Message}");
                throw;
            }
        }

        public void RegisterMessageHandler(string messageType, Action<JsonElement> handler)
        {
            if (_messageHandlers.ContainsKey(messageType))
            {
                _messageHandlers[messageType] = handler;
            }
            else
            {
                _messageHandlers.Add(messageType, handler);
            }
        }

        private async Task HandleWebSocketConnectionAsync(WebSocket webSocket)
        {
            _activeConnections.Add(webSocket);
            var buffer = new byte[1024 * 4];
            try
            {
                Log("새로운 웹소켓 클라이언트 연결 요청이 들어왔습니다.");
                Log($"현재 활성 연결 수: {_activeConnections.Count}");
                MessageDispatcher.Instance.DispatchToUI(() => OnConnected?.Invoke());

                // 연결 성공 메시지 전송
                var response = new { status = "connected", message = "서버에 연결되었습니다." };
                var responseBytes = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(response));
                Log("클라이언트에게 연결 성공 메시지를 전송합니다.");
                await webSocket.SendAsync(new ArraySegment<byte>(responseBytes), WebSocketMessageType.Text, true, CancellationToken.None);
                Log("연결 성공 메시지 전송 완료");

                while (webSocket.State == WebSocketState.Open)
                {
                    try
                    {
                        Log("클라이언트로부터 메시지 수신 대기 중...");
                        var result = await webSocket.ReceiveAsync(new ArraySegment<byte>(buffer), CancellationToken.None);
                        Log($"메시지 수신 완료: 타입={result.MessageType}, 길이={result.Count}바이트");
                        
                        if (result.MessageType == WebSocketMessageType.Text)
                        {
                            var message = Encoding.UTF8.GetString(buffer, 0, result.Count);
                            Log($"메시지 내용: {message}");
                            
                            // JSON 메시지 파싱 및 처리
                            var jsonDoc = JsonDocument.Parse(message);
                            var root = jsonDoc.RootElement;
                            
                            // command 또는 type 필드 확인
                            string? messageType = null;
                            if (root.TryGetProperty("command", out var commandElement))
                            {
                                messageType = commandElement.GetString();
                                Log($"명령어 타입: {messageType}");
                            }
                            else if (root.TryGetProperty("type", out var typeElement))
                            {
                                messageType = typeElement.GetString();
                                Log($"메시지 타입: {messageType}");
                            }

                            if (!string.IsNullOrEmpty(messageType))
                            {
                                if (_messageHandlers.TryGetValue(messageType, out var handler))
                                {
                                    Log($"메시지 핸들러 실행: {messageType}");
                                    handler(root);
                                }
                                Log("메시지 디스패처로 전달");
                                MessageDispatcher.Instance.DispatchJsonMessage(messageType, root);
                            }
                            else
                            {
                                Log("메시지에 command 또는 type 필드가 없습니다.");
                            }

                            // 메시지 수신 확인 응답 전송
                            var ackResponse = new { status = "received", message = "메시지를 받았습니다." };
                            var ackBytes = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(ackResponse));
                            Log("메시지 수신 확인 응답 전송");
                            await webSocket.SendAsync(new ArraySegment<byte>(ackBytes), WebSocketMessageType.Text, true, CancellationToken.None);
                            Log("메시지 수신 확인 응답 전송 완료");
                        }
                        else if (result.MessageType == WebSocketMessageType.Close)
                        {
                            Log("클라이언트가 연결 종료를 요청했습니다.");
                            await webSocket.CloseAsync(WebSocketCloseStatus.NormalClosure, string.Empty, CancellationToken.None);
                            Log("연결이 정상적으로 종료되었습니다.");
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"메시지 처리 중 오류 발생: {ex.Message}");
                        Log($"스택 트레이스: {ex.StackTrace}");
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"웹소켓 연결 처리 중 오류 발생: {ex.Message}");
                Log($"스택 트레이스: {ex.StackTrace}");
            }
            finally
            {
                _activeConnections.Remove(webSocket);
                Log($"연결이 제거되었습니다. 현재 활성 연결 수: {_activeConnections.Count}");
                if (webSocket.State == WebSocketState.Open)
                {
                    try
                    {
                        Log("웹소켓 연결을 정상적으로 종료합니다.");
                        await webSocket.CloseAsync(WebSocketCloseStatus.NormalClosure, string.Empty, CancellationToken.None);
                        Log("웹소켓 연결 종료 완료");
                    }
                    catch (Exception ex)
                    {
                        Log($"웹소켓 종료 중 오류 발생: {ex.Message}");
                        Log($"스택 트레이스: {ex.StackTrace}");
                    }
                }
            }
        }

        public async Task SendMessageAsync(string command, object parameters)
        {
            try
            {
                var message = new
                {
                    command = command,
                    parameters = parameters
                };

                var jsonMessage = JsonSerializer.Serialize(message);
                var buffer = Encoding.UTF8.GetBytes(jsonMessage);

                foreach (var webSocket in _activeConnections)
                {
                    if (webSocket.State == WebSocketState.Open)
                    {
                        await webSocket.SendAsync(
                            new ArraySegment<byte>(buffer),
                            WebSocketMessageType.Text,
                            true,
                            CancellationToken.None
                        );
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"메시지 전송 중 오류 발생: {ex.Message}");
                throw;
            }
        }

        public async Task StopAsync()
        {
            if (!_isRunning)
            {
                return;
            }

            try
            {
                _cancellationTokenSource?.Cancel();

                // 모든 활성 연결 종료
                foreach (var webSocket in _activeConnections.ToList())
                {
                    try
                    {
                        if (webSocket.State == WebSocketState.Open)
                        {
                            await webSocket.CloseAsync(WebSocketCloseStatus.NormalClosure, "서버 종료", CancellationToken.None);
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"웹소켓 연결 종료 중 오류 발생: {ex.Message}");
                    }
                }
                _activeConnections.Clear();

                _httpListener?.Stop();
                _isRunning = false;
                Log("웹소켓 서버가 중지되었습니다.");
            }
            catch (Exception ex)
            {
                Log($"서버 종료 중 오류 발생: {ex.Message}");
                throw;
            }
        }
    }
} 