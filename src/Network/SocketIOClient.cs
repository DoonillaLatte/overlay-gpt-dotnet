using System;
using System.Threading;
using System.Threading.Tasks;
using SocketIOClient;
using System.Text.Json;
using SocketIO.Core;

namespace overlay_gpt.Network
{
    public class SocketIOConnection
    {
        private SocketIOClient.SocketIO _socket;
        private readonly string _serverUrl;
        private readonly CancellationTokenSource _cancellationTokenSource;

        public event EventHandler<string>? OnMessageReceived;

        public SocketIOConnection(string serverUrl = "http://localhost:5000")
        {
            _serverUrl = serverUrl;
            _cancellationTokenSource = new CancellationTokenSource();
            
            var options = new SocketIOClient.SocketIOOptions
            {
                EIO = EngineIO.V4,
                Transport = SocketIOClient.Transport.TransportProtocol.WebSocket,
                Reconnection = true,
                ReconnectionAttempts = 5,
                ReconnectionDelay = 1000,
                ReconnectionDelayMax = 5000
            };

            _socket = new SocketIOClient.SocketIO(_serverUrl, options);
            
            _socket.On("message_response", response =>
            {
                Console.WriteLine("==========================================");
                Console.WriteLine("message_response 이벤트 수신됨");
                Console.WriteLine($"수신 시간: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                Console.WriteLine($"이벤트 타입: message_response");
                Console.WriteLine($"전체 응답 데이터: {response}");
                var message = response.GetValue<string>();
                Console.WriteLine($"파싱된 메시지: {message}");
                Console.WriteLine("==========================================");
                OnMessageReceived?.Invoke(this, message);
            });

            _socket.On("prompt_response", response =>
            {
                Console.WriteLine("==========================================");
                Console.WriteLine("prompt_response 이벤트 수신됨");
                Console.WriteLine($"수신 시간: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                Console.WriteLine($"이벤트 타입: prompt_response");
                Console.WriteLine($"전체 응답 데이터: {response}");
                var data = response.GetValue<object>();
                Console.WriteLine($"파싱된 데이터: {data}");
                Console.WriteLine("==========================================");
                OnMessageReceived?.Invoke(this, JsonSerializer.Serialize(data));
            });

            _socket.OnConnected += (sender, e) =>
            {
                Console.WriteLine("==========================================");
                Console.WriteLine("Socket.IO 서버에 연결되었습니다.");
                Console.WriteLine($"연결 시간: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                Console.WriteLine("==========================================");
                LogWindow.Instance.Log("Socket.IO 서버에 연결되었습니다.");
            };

            _socket.OnDisconnected += (sender, e) =>
            {
                Console.WriteLine("==========================================");
                Console.WriteLine("Socket.IO 서버와 연결이 끊어졌습니다.");
                Console.WriteLine($"연결 해제 시간: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                Console.WriteLine("==========================================");
                LogWindow.Instance.Log("Socket.IO 서버와 연결이 끊어졌습니다.");
            };

            _socket.OnError += (sender, e) =>
            {
                Console.WriteLine("==========================================");
                Console.WriteLine($"Socket.IO 오류 발생: {e}");
                Console.WriteLine($"오류 발생 시간: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                Console.WriteLine("==========================================");
                LogWindow.Instance.Log($"Socket.IO 오류 발생: {e}");
            };
        }

        public async Task ConnectAsync()
        {
            try
            {
                await _socket.ConnectAsync();
            }
            catch (Exception ex)
            {
                LogWindow.Instance.Log($"Socket.IO 연결 실패: {ex.Message}");
                throw;
            }
        }

        public async Task SendMessageAsync(string message)
        {
            if (!_socket.Connected)
            {
                throw new InvalidOperationException("Socket.IO가 연결되어 있지 않습니다.");
            }

            try
            {
                await _socket.EmitAsync("message", message);
            }
            catch (Exception ex)
            {
                LogWindow.Instance.Log($"메시지 전송 실패: {ex.Message}");
                throw;
            }
        }

        public async Task DisconnectAsync()
        {
            if (_socket.Connected)
            {
                await _socket.DisconnectAsync();
            }
            _cancellationTokenSource.Cancel();
            _socket.Dispose();
        }
    }
} 