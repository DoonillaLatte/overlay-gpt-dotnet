using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Microsoft.AspNetCore.SignalR;
using overlay_gpt.Network.Models.Vue;
using overlay_gpt.Network.Models.Flask;
using overlay_gpt.Services;
using overlay_gpt.Network.Models.Common;
using System.Text.Json;

namespace overlay_gpt.Network
{
    public class ProcessVueMessage
    {
        private readonly Dictionary<string, Func<JObject, Task>> _commandHandlers;
        private readonly IHubContext<ChatHub> _hubContext;
        private readonly WithFlaskAsClient _flaskClient;

        public ProcessVueMessage(IHubContext<ChatHub> hubContext, WithFlaskAsClient flaskClient)
        {
            _hubContext = hubContext;
            _flaskClient = flaskClient;
            _commandHandlers = new Dictionary<string, Func<JObject, Task>>
            {
                { "show_overlay", HandleShowOverlay },
                { "hide_overlay", HandleHideOverlay },
                { "update_content", HandleUpdateContent },
                { "error", HandleError },
                { "send_user_prompt", HandleSendUserPrompt },
                { "ping", HandlePing },
                { "generate_chat_id", HandleGenerateChatId }
            };
        }

        public async Task ProcessMessage(string connectionId, JObject message)
        {
            try
            {
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 메시지 수신 시작");
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ConnectionId: {connectionId}");
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 메시지 내용: {message.ToString(Newtonsoft.Json.Formatting.Indented)}");
                
                var command = message["command"]?.ToString();
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 명령어: {command}");

                if (string.IsNullOrEmpty(command))
                {
                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 오류: 명령어가 지정되지 않았습니다.");
                    await SendErrorResponse(connectionId, "명령어가 지정되지 않았습니다.");
                    return;
                }

                if (_commandHandlers.TryGetValue(command, out var handler))
                {
                    await handler(message);
                }
                else
                {
                    Console.WriteLine($"처리되지 않은 명령어: {command}");
                    await SendErrorResponse(connectionId, $"처리되지 않은 명령어: {command}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"메시지 처리 중 오류 발생: {ex.Message}");
                await SendErrorResponse(connectionId, $"메시지 처리 중 오류 발생: {ex.Message}");
            }
        }

        private async Task HandleShowOverlay(JObject data)
        {
            Console.WriteLine("오버레이를 표시합니다.");
            // TODO: 실제 오버레이 표시 로직 구현
            await Task.CompletedTask;
        }

        private async Task HandleHideOverlay(JObject data)
        {
            Console.WriteLine("오버레이를 숨깁니다.");
            // TODO: 실제 오버레이 숨기기 로직 구현
            await Task.CompletedTask;
        }

        private async Task HandleUpdateContent(JObject data)
        {
            var content = data["content"]?.ToString();
            Console.WriteLine($"콘텐츠 업데이트: {content}");
            // TODO: 실제 콘텐츠 업데이트 로직 구현
            await Task.CompletedTask;
        }

        private async Task HandleError(JObject data)
        {
            var errorMessage = data["message"]?.ToString();
            Console.WriteLine($"에러 발생: {errorMessage}");
            // TODO: 실제 에러 처리 로직 구현
            await Task.CompletedTask;
        }

        private async Task HandleSendUserPrompt(JObject data)
        {
            try
            {
                var vueRequest = data.ToObject<VueRequest>();
                if (vueRequest == null)
                {
                    throw new Exception("잘못된 요청 형식입니다.");
                }

                var chatData = ChatDataManager.Instance.GetChatDataById(vueRequest.ChatId);
                if (chatData == null)
                {
                    chatData = new ChatData
                    {
                        ChatId = vueRequest.ChatId,
                        GeneratedTimestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                        CurrentProgram = vueRequest.CurrentProgram,
                        TargetProgram = vueRequest.TargetProgram
                    };
                    ChatDataManager.Instance.AddChatData(chatData);
                    Console.WriteLine($"해당 ChatData가 없어 새로 생성합니다. ID : {vueRequest.ChatId}");
                }

                var flaskRequest = new RequestPrompt
                {
                    ChatId = vueRequest.ChatId,
                    Prompt = vueRequest.Prompt,
                    GeneratedTimestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    RequestType = 1,
                    CurrentProgram = vueRequest.CurrentProgram,
                    TargetProgram = vueRequest.TargetProgram
                };

                await _flaskClient.EmitAsync("message", flaskRequest);
                Console.WriteLine("Flask 서버로 메시지를 전송했습니다.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"메시지 변환 중 오류 발생: {ex.Message}");
                throw;
            }
        }

        private async Task HandlePing(JObject data)
        {
            try
            {
                var response = new
                {
                    status = "success",
                    message = "pong",
                    timestamp = DateTime.Now
                };
                await _hubContext.Clients.Client(data["connectionId"]?.ToString()).SendAsync("ReceiveMessage", response);
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Ping 요청에 Pong 응답을 보냈습니다.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ping 처리 중 오류 발생: {ex.Message}");
                await SendErrorResponse(data["connectionId"]?.ToString(), $"Ping 처리 중 오류 발생: {ex.Message}");
            }
        }

        private async Task HandleGenerateChatId(JObject data)
        {
            try
            {
                var chatId = data["chat_id"]?.Value<int>();
                var generatedTimestamp = data["generated_timestamp"]?.ToString();

                if (chatId == null || string.IsNullOrEmpty(generatedTimestamp))
                {
                    throw new Exception("chat_id 또는 generated_timestamp가 누락되었습니다.");
                }

                var chatData = ChatDataManager.Instance.GetChatDataByTimeStamp(generatedTimestamp);
                if (chatData != null)
                {
                    chatData.ChatId = chatId.Value;
                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ChatData ID 업데이트 완료: {chatId} (Timestamp: {generatedTimestamp})");
                }
                else
                {
                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 해당 Timestamp({generatedTimestamp})에 일치하는 ChatData를 찾을 수 없습니다.");
                }

                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Chat ID 생성 처리 중 오류 발생: {ex.Message}");
                throw;
            }
        }

        private async Task SendErrorResponse(string connectionId, string errorMessage)
        {
            var response = new
            {
                status = "error",
                message = errorMessage
            };
            await _hubContext.Clients.Client(connectionId).SendAsync("ReceiveMessage", response);
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
