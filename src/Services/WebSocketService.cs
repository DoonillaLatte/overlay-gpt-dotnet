using System;
using System.Net.WebSockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Text.Json;
using System.Collections.Generic;

namespace overlay_gpt.Services
{
    public class WebSocketService
    {
        private WebSocket? _webSocket;
        private readonly CancellationTokenSource _cancellationTokenSource;
        private const int ReceiveBufferSize = 8192;
        private readonly Dictionary<string, Func<JsonElement, Task>> _commandHandlers;

        public WebSocketService()
        {
            _cancellationTokenSource = new CancellationTokenSource();
            _commandHandlers = new Dictionary<string, Func<JsonElement, Task>>();
        }

        public void RegisterCommandHandler(string commandName, Func<JsonElement, Task> handler)
        {
            _commandHandlers[commandName] = handler;
        }

        public async Task StartAsync(string url)
        {
            try
            {
                using var client = new ClientWebSocket();
                _webSocket = client;
                await client.ConnectAsync(new Uri(url), _cancellationTokenSource.Token);
                Console.WriteLine("웹소켓 연결이 성공적으로 설정되었습니다.");

                // 메시지 수신 루프 시작
                _ = ReceiveMessagesAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"웹소켓 연결 중 오류 발생: {ex.Message}");
                throw;
            }
        }

        private async Task ReceiveMessagesAsync()
        {
            var buffer = new byte[ReceiveBufferSize];
            try
            {
                while (_webSocket?.State == WebSocketState.Open)
                {
                    var result = await _webSocket.ReceiveAsync(
                        new ArraySegment<byte>(buffer), _cancellationTokenSource.Token);

                    if (result.MessageType == WebSocketMessageType.Close)
                    {
                        await _webSocket.CloseAsync(
                            WebSocketCloseStatus.NormalClosure,
                            "클라이언트에 의해 연결이 종료됨",
                            _cancellationTokenSource.Token);
                        break;
                    }

                    var message = Encoding.UTF8.GetString(buffer, 0, result.Count);
                    await HandleMessageAsync(message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"메시지 수신 중 오류 발생: {ex.Message}");
            }
        }

        public async Task SendCommandAsync(string commandName, object parameters)
        {
            if (_webSocket?.State != WebSocketState.Open)
            {
                throw new InvalidOperationException("웹소켓이 연결되어 있지 않습니다.");
            }

            var message = new
            {
                command = commandName,
                parameters = parameters
            };

            var json = JsonSerializer.Serialize(message);
            var buffer = Encoding.UTF8.GetBytes(json);
            
            await _webSocket.SendAsync(
                new ArraySegment<byte>(buffer),
                WebSocketMessageType.Text,
                true,
                _cancellationTokenSource.Token);
        }

        private async Task HandleMessageAsync(string message)
        {
            try
            {
                using var jsonDoc = JsonDocument.Parse(message);
                var root = jsonDoc.RootElement;

                if (!root.TryGetProperty("command", out var commandElement) ||
                    !root.TryGetProperty("parameters", out var parametersElement))
                {
                    Console.WriteLine("잘못된 메시지 형식: command와 parameters가 필요합니다.");
                    return;
                }

                var commandName = commandElement.GetString();
                if (string.IsNullOrEmpty(commandName))
                {
                    Console.WriteLine("command 이름이 비어있습니다.");
                    return;
                }

                if (_commandHandlers.TryGetValue(commandName, out var handler))
                {
                    await handler(parametersElement);
                }
                else
                {
                    Console.WriteLine($"처리되지 않은 command: {commandName}");
                }
            }
            catch (JsonException ex)
            {
                Console.WriteLine($"JSON 파싱 오류: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"메시지 처리 중 오류 발생: {ex.Message}");
            }
        }

        public async Task StopAsync()
        {
            if (_webSocket?.State == WebSocketState.Open)
            {
                await _webSocket.CloseAsync(
                    WebSocketCloseStatus.NormalClosure,
                    "서비스 종료",
                    _cancellationTokenSource.Token);
            }
            _cancellationTokenSource.Cancel();
        }
    }
} 