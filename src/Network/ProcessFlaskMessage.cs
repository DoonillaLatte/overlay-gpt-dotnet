using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using SocketIOClient;
using Newtonsoft.Json.Linq;
using overlay_gpt.Network.Models.Common;
using overlay_gpt.Network.Models.Vue;

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
                { "error", HandleError },
                { "generated_response", HandleGeneratedResponse },
                { "response_workflows", HandleResponseWorkflows }
            };
        }

        public async Task ProcessMessage(SocketIOResponse response)
        {
            try
            {
                Console.WriteLine("ProcessMessage 시작");
                var jsonString = response.GetValue<object>().ToString();
                var jsonData = JObject.Parse(jsonString);
                Console.WriteLine($"수신된 메시지: {jsonData}");
                
                var command = jsonData["command"]?.ToString();
                Console.WriteLine($"명령어: {command}");

                if (string.IsNullOrEmpty(command))
                {
                    Console.WriteLine("명령어가 지정되지 않았습니다.");
                    return;
                }

                Console.WriteLine($"사용 가능한 핸들러 목록: {string.Join(", ", _commandHandlers.Keys)}");
                
                if (_commandHandlers.TryGetValue(command, out var handler))
                {
                    Console.WriteLine($"핸들러 실행: {command}");
                    await handler(jsonData);
                    Console.WriteLine($"핸들러 실행 완료: {command}");
                }
                else
                {
                    Console.WriteLine($"처리되지 않은 명령어: {command}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"메시지 처리 중 오류 발생: {ex.Message}");
                Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
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

        private async Task HandleGeneratedResponse(JObject data)
        {
            try
            {
                Console.WriteLine("HandleGeneratedResponse 시작");
                var chatId = data["chat_id"]?.Value<int>() ?? -1;
                var title = data["title"]?.ToString();
                var message = data["message"]?.ToString();
                var status = data["status"]?.ToString();

                Console.WriteLine($"받은 데이터 - chatId: {chatId}, message: {message}, status: {status}");

                if (string.IsNullOrEmpty(message))
                {
                    Console.WriteLine("message가 비어있습니다.");
                    return;
                }

                // HTML 형식인지 확인
                bool isHtml = message.Contains("<") && message.Contains(">");
                string textType = isHtml ? "text_to_apply" : "text_plain";

                var chatData = Services.ChatDataManager.Instance.GetChatDataById(chatId);
                if (chatData == null)
                {
                    Console.WriteLine($"chat_id {chatId}에 해당하는 ChatData를 찾을 수 없습니다.");
                    return;
                }

                if(chatData.Title != title) 
                {
                    chatData.Title = title;
                }
                chatData.Texts.Add(new TextData { Type = textType, Content = message });
                if(chatData.TargetProgram == null) 
                {
                    chatData.CurrentProgram.Context = message;
                }
                else
                {
                    chatData.TargetProgram.Context = message;
                }
                Console.WriteLine($"ChatData {chatId}에 메시지가 추가되었습니다.");

                // Vue로 display_text 메시지 전송
                var displayText = new DisplayText
                {
                    ChatId = chatId,
                    Title = title,
                    GeneratedTimestamp = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                    CurrentProgram = chatData.CurrentProgram,
                    TargetProgram = chatData.TargetProgram,
                    Texts = new List<TextData> { new TextData { Type = textType, Content = message } }
                };

                var vueServer = MainWindow.Instance.VueServer;
                if (vueServer != null)
                {
                    await vueServer.SendMessageToAll(displayText);
                    Console.WriteLine($"Vue로 display_text 메시지 전송 완료: chat_id {chatId}");
                }
                else
                {
                    Console.WriteLine("Vue 서버가 초기화되지 않았습니다.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"메시지 처리 중 오류 발생: {ex.Message}");
            }
        }

        private async Task HandleResponseWorkflows(JObject data)
        {
            try
            {
                Console.WriteLine("HandleResponseWorkflows 시작");
                // TODO: 워크플로우 응답 처리 로직 구현
            }
            catch (Exception ex)
            {
                Console.WriteLine($"워크플로우 응답 처리 중 오류 발생: {ex.Message}");
            }
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
