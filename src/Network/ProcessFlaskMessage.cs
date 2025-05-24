using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using SocketIOClient;
using Newtonsoft.Json.Linq;

namespace overlay_gpt.Network
{
    public class ProcessFlaskMessage
    {
        private readonly Dictionary<string, Func<JObject, Task>> _commandHandlers;

        public ProcessFlaskMessage()
        {
            _commandHandlers = new Dictionary<string, Func<JObject, Task>>
            {
                { "show_overlay", HandleShowOverlay },
                { "hide_overlay", HandleHideOverlay },
                { "update_content", HandleUpdateContent },
                { "error", HandleError }
            };
        }

        public async Task ProcessMessage(SocketIOResponse response)
        {
            try
            {
                var jsonData = response.GetValue<JObject>();
                Console.WriteLine($"수신된 메시지: {jsonData}");
                
                var command = jsonData["command"]?.ToString();

                if (string.IsNullOrEmpty(command))
                {
                    Console.WriteLine("명령어가 지정되지 않았습니다.");
                    return;
                }

                if (_commandHandlers.TryGetValue(command, out var handler))
                {
                    await handler(jsonData);
                }
                else
                {
                    Console.WriteLine($"처리되지 않은 명령어: {command}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"메시지 처리 중 오류 발생: {ex.Message}");
            }
        }

        private async Task HandleShowOverlay(JObject data)
        {
            // 오버레이 표시 로직 구현
            Console.WriteLine("오버레이를 표시합니다.");
            // TODO: 실제 오버레이 표시 로직 구현
        }

        private async Task HandleHideOverlay(JObject data)
        {
            // 오버레이 숨기기 로직 구현
            Console.WriteLine("오버레이를 숨깁니다.");
            // TODO: 실제 오버레이 숨기기 로직 구현
        }

        private async Task HandleUpdateContent(JObject data)
        {
            var content = data["content"]?.ToString();
            Console.WriteLine($"콘텐츠 업데이트: {content}");
            // TODO: 실제 콘텐츠 업데이트 로직 구현
        }

        private async Task HandleError(JObject data)
        {
            var errorMessage = data["message"]?.ToString();
            Console.WriteLine($"에러 발생: {errorMessage}");
            // TODO: 실제 에러 처리 로직 구현
        }

        // 새로운 명령어 핸들러를 추가하는 메서드
        public void RegisterCommandHandler(string command, Func<JObject, Task> handler)
        {
            if (_commandHandlers.ContainsKey(command))
            {
                _commandHandlers[command] = handler;
            }
            else
            {
                _commandHandlers.Add(command, handler);
            }
        }
    }
}
