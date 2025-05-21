using System;
using System.Threading;
using System.Threading.Tasks;
using SocketIOClient;
using System.Text.Json;
using SocketIO.Core;
using overlay_gpt.Network.Models;
using Microsoft.AspNetCore.SignalR;

namespace overlay_gpt.Network
{
    public class SocketIOConnection
    {
        private SocketIOClient.SocketIO _socket;
        private readonly string _serverUrl;
        private readonly CancellationTokenSource _cancellationTokenSource;
        private readonly IHubContext<ChatHub> _hubContext;

        public event EventHandler<string>? OnMessageReceived;

        public SocketIOConnection(IHubContext<ChatHub> hubContext, string serverUrl = "http://localhost:5000")
        {
            _serverUrl = serverUrl;
            _cancellationTokenSource = new CancellationTokenSource();
            _hubContext = hubContext;
            
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
            
            _socket.On("message_response", async response =>
            {
                Console.WriteLine("==========================================");
                Console.WriteLine("message_response 이벤트 수신됨");
                Console.WriteLine($"수신 시간: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                Console.WriteLine($"이벤트 타입: message_response");
                Console.WriteLine($"전체 응답 데이터: {response}");
                
                try
                {
                    var jsonElement = response.GetValue<JsonElement>();
                    Console.WriteLine($"파싱된 JSON: {JsonSerializer.Serialize(jsonElement, new JsonSerializerOptions { WriteIndented = true })}");
                    
                    if (jsonElement.TryGetProperty("command", out var commandElement))
                    {
                        var command = commandElement.GetString();
                        switch (command)
                        {
                            case "response_single_generated_response":
                                Console.WriteLine("response_single_generated_response 처리 시작");
                                Console.WriteLine($"JSON 데이터 구조: {JsonSerializer.Serialize(jsonElement, new JsonSerializerOptions { WriteIndented = true })}");
                                
                                if (!jsonElement.TryGetProperty("message", out var responseElement) ||
                                    !jsonElement.TryGetProperty("status", out var statusElement))
                                {
                                    Console.WriteLine("필수 필드 누락 - message 또는 status가 없습니다.");
                                    throw new InvalidOperationException("필수 필드가 누락되었습니다: message 또는 status");
                                }

                                int responseChatId = jsonElement.TryGetProperty("chat_id", out var responseChatIdElement) 
                                    ? responseChatIdElement.GetInt32() 
                                    : 1; // 임시
                                Console.WriteLine($"chat_id 처리 결과: {responseChatId}");

                                string responseText = responseElement.GetString() ?? throw new InvalidOperationException("message 값이 null입니다.");
                                string responseStatus = statusElement.GetString() ?? throw new InvalidOperationException("status 값이 null입니다.");
                                Console.WriteLine($"message 처리 결과: {responseText}");
                                Console.WriteLine($"status 처리 결과: {responseStatus}");

                                Console.WriteLine("current_program 정보 추출 중...");
                                var currentProgram = new
                                {
                                    file_name = jsonElement.TryGetProperty("current_program", out var currProgElement) 
                                        ? currProgElement.GetProperty("file_name").GetString() ?? string.Empty
                                        : string.Empty,
                                    program_type = jsonElement.TryGetProperty("current_program", out var currProgElement2) 
                                        ? currProgElement2.GetProperty("program_type").GetString() ?? string.Empty
                                        : string.Empty,
                                    context = jsonElement.TryGetProperty("current_program", out var currProgElement3) 
                                        ? currProgElement3.GetProperty("context").GetString() ?? string.Empty
                                        : string.Empty
                                };
                                Console.WriteLine($"current_program: {JsonSerializer.Serialize(currentProgram)}");

                                Console.WriteLine("target_program 정보 추출 중...");
                                var targetProgram = jsonElement.TryGetProperty("target_program", out var targetProgElement) 
                                    ? new
                                    {
                                        file_name = targetProgElement.GetProperty("file_name").GetString() ?? string.Empty,
                                        program_type = targetProgElement.GetProperty("program_type").GetString() ?? string.Empty,
                                        context = targetProgElement.GetProperty("context").GetString() ?? string.Empty
                                    }
                                    : null;
                                Console.WriteLine($"target_program: {JsonSerializer.Serialize(targetProgram)}");

                                var displayTextMessage = new
                                {
                                    command = "display_text",
                                    chat_id = responseChatId,
                                    current_program = currentProgram,
                                    target_program = targetProgram,
                                    texts = new[]
                                    {
                                        new
                                        {
                                            type = "text_plain",
                                            content = responseText
                                        }
                                    }
                                };

                                Console.WriteLine("최종 메시지 생성 완료:");
                                Console.WriteLine(JsonSerializer.Serialize(displayTextMessage, new JsonSerializerOptions { WriteIndented = true }));
                                
                                // SignalR을 통해 Vue로 메시지 전송
                                Console.WriteLine("==========================================");
                                Console.WriteLine("Vue로 메시지 전송 중...");
                                Console.WriteLine($"전송 시간: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                                Console.WriteLine($"전송할 메시지: {JsonSerializer.Serialize(displayTextMessage, new JsonSerializerOptions { WriteIndented = true })}");
                                await _hubContext.Clients.All.SendAsync("ReceiveMessage", displayTextMessage);
                                Console.WriteLine("Vue로 메시지 전송 완료");
                                Console.WriteLine("==========================================");
                                
                                OnMessageReceived?.Invoke(this, JsonSerializer.Serialize(displayTextMessage));
                                break;
                            default:
                                Console.WriteLine($"알 수 없는 명령: {command}");
                                Console.WriteLine("==========================================");
                                Console.WriteLine("Vue로 메시지 전송 중...");
                                Console.WriteLine($"전송 시간: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                                Console.WriteLine($"전송할 메시지: {response}");
                                await _hubContext.Clients.All.SendAsync("ReceiveMessage", response.ToString());
                                Console.WriteLine("Vue로 메시지 전송 완료");
                                Console.WriteLine("==========================================");
                                OnMessageReceived?.Invoke(this, response.ToString());
                                break;
                        }
                    }
                }
                catch (JsonException ex)
                {
                    Console.WriteLine($"JSON 파싱 오류: {ex.Message}");
                    Console.WriteLine("==========================================");
                    Console.WriteLine("Vue로 메시지 전송 중...");
                    Console.WriteLine($"전송 시간: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    Console.WriteLine($"전송할 메시지: {response}");
                    await _hubContext.Clients.All.SendAsync("ReceiveMessage", response.ToString());
                    Console.WriteLine("Vue로 메시지 전송 완료");
                    Console.WriteLine("==========================================");
                    OnMessageReceived?.Invoke(this, response.ToString());
                }
                
                Console.WriteLine("==========================================");
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