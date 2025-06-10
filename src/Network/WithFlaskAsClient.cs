/*
 * Flask SocketIO 서버와 통신하는 클라이언트
 * 
 * 사용 방법:
 * 1. 클라이언트 인스턴스 생성
 *    var client = new WithFlaskAsClient();
 * 
 * 2. 서버 연결
 *    await client.ConnectAsync();
 * 
 * 3. 메시지 전송 방법
 *    // 문자열 메시지 전송
 *    await client.SendMessageAsync("안녕하세요!");
 * 
 *    // 복잡한 데이터 전송
 *    await client.SendMessageAsync(new { 
 *        type = "notification",
 *        content = "새로운 메시지가 도착했습니다",
 *        timestamp = DateTime.Now
 *    });
 * 
 * 4. 서버로부터 메시지 수신
 *    client.On("message", (response) => {
 *        Console.WriteLine($"수신된 메시지: {response.GetValue<string>()}");
 *    });
 * 
 * 5. 연결 종료
 *    await client.DisconnectAsync();
 */

using System;
using System.Threading.Tasks;
using SocketIOClient;
using SocketIOClient.Transport;

namespace overlay_gpt.Network
{
    public class WithFlaskAsClient
    {
        private SocketIOClient.SocketIO _socket;
        private const string ServerUrl = "http://localhost:5001";
        private const string MessageEvent = "message";
        private const string MessageResponseEvent = "message_response";
        private const int ReconnectionDelay = 2000;
        private int _reconnectionAttempt = 0;
        private readonly ProcessFlaskMessage _messageProcessor;

        public WithFlaskAsClient()
        {
            _messageProcessor = new ProcessFlaskMessage();
        }

        public async Task ConnectAsync()
        {
            _socket = new SocketIOClient.SocketIO(ServerUrl, new SocketIOOptions
            {
                Transport = TransportProtocol.WebSocket,
                Reconnection = true,
                ReconnectionAttempts = int.MaxValue,
                ReconnectionDelay = ReconnectionDelay
            });

            _socket.OnConnected += (sender, e) =>
            {
                Console.WriteLine("서버에 연결되었습니다.");
                _reconnectionAttempt = 0;
            };

            _socket.OnDisconnected += (sender, e) =>
            {
                Console.WriteLine("서버와 연결이 끊어졌습니다.");
                _reconnectionAttempt++;
                
                Console.WriteLine($"{ReconnectionDelay/1000}초 후 재연결을 시도합니다. (시도 {_reconnectionAttempt})");
                Task.Delay(ReconnectionDelay).ContinueWith(_ => ConnectAsync());
            };

            _socket.OnError += (sender, e) =>
            {
                Console.WriteLine($"에러 발생: {e}");
            };

            _socket.Off(MessageResponseEvent);
            _socket.On(MessageResponseEvent, async (response) =>
            {
                Console.WriteLine($"message_response 이벤트 수신: {response.GetValue<object>()}");
                await _messageProcessor.ProcessMessage(response);
            });

            await _socket.ConnectAsync();
        }

        public async Task DisconnectAsync()
        {
            if (_socket != null)
            {
                await _socket.DisconnectAsync();
            }
        }

        public async Task EmitAsync(string eventName, object data)
        {
            if (_socket != null && _socket.Connected)
            {
                await _socket.EmitAsync(eventName, data);
            }
        }

        public void On(string eventName, Action<SocketIOResponse> callback)
        {
            if (_socket != null)
            {
                _socket.On(eventName, callback);
            }
        }

        public async Task SendMessageAsync(string message)
        {
            if (_socket == null || !_socket.Connected)
            {
                throw new InvalidOperationException("서버에 연결되어 있지 않습니다.");
            }

            try
            {
                await _socket.EmitAsync(MessageEvent, new { text = message });
                Console.WriteLine($"메시지 전송 완료: {message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"메시지 전송 실패: {ex.Message}");
                throw;
            }
        }

        public async Task SendMessageAsync(object messageData)
        {
            if (_socket == null || !_socket.Connected)
            {
                throw new InvalidOperationException("서버에 연결되어 있지 않습니다.");
            }

            try
            {
                Console.WriteLine($"메시지 전송 시작: {messageData}");
                await _socket.EmitAsync(MessageEvent, messageData);
                Console.WriteLine($"메시지 전송 완료: {messageData}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"메시지 전송 실패: {ex.Message}");
                throw;
            }
        }
    }
}
