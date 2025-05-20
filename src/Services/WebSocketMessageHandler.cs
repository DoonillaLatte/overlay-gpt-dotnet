using System;
using System.Text.Json;
using System.Threading.Tasks;
using overlay_gpt.Models;
using overlay_gpt.Services;

namespace overlay_gpt.Services
{
    public class WebSocketMessageHandler
    {
        private readonly Dictionary<string, Func<JsonElement, Task>> _messageHandlers;
        private readonly LogWindow _logWindow;

        public WebSocketMessageHandler(LogWindow logWindow)
        {
            _messageHandlers = new Dictionary<string, Func<JsonElement, Task>>();
            _logWindow = logWindow;
            RegisterDefaultHandlers();
        }

        private void RegisterDefaultHandlers()
        {
            RegisterHandler("send_user_prompt", HandleUserPromptAsync);
            RegisterHandler("chat", HandleChatMessageAsync);
        }

        private async Task HandleChatMessageAsync(JsonElement json)
        {
            try
            {
                var message = json.GetProperty("parameters").GetProperty("message").GetString();
                _logWindow.Log($"채팅 메시지 수신: {message}");
                
                // 여기에 채팅 메시지 처리 로직 추가
                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                _logWindow.Log($"채팅 메시지 처리 중 오류 발생: {ex.Message}");
            }
        }

        private async Task HandleUserPromptAsync(JsonElement json)
        {
            try
            {
                var userPrompt = await Task.Run(() => JsonSerializer.Deserialize<UserPrompt>(json.GetProperty("parameters").GetRawText()));
                if (userPrompt == null)
                {
                    _logWindow.Log("사용자 프롬프트 파싱 실패");
                    return;
                }

                // TODO: 여기에 실제 처리 로직 구현
                await Task.Run(() => {
                    _logWindow.Log($"채팅 ID: {userPrompt.ChatId}");
                    _logWindow.Log($"프롬프트: {userPrompt.Prompt}");
                    _logWindow.Log($"대상 프로그램: {userPrompt.TargetProgram ?? "없음"}");
                });
            }
            catch (Exception ex)
            {
                _logWindow.Log($"사용자 프롬프트 처리 중 오류 발생: {ex.Message}");
            }
        }

        public void RegisterHandler(string messageType, Func<JsonElement, Task> handler)
        {
            _messageHandlers[messageType] = handler;
        }

        public async Task HandleMessageAsync(string message)
        {
            try
            {
                var jsonDocument = JsonDocument.Parse(message);
                var root = jsonDocument.RootElement;

                if (root.TryGetProperty("command", out var commandElement))
                {
                    var command = commandElement.GetString();
                    if (command != null && _messageHandlers.TryGetValue(command, out var handler))
                    {
                        await handler(root);
                    }
                    else
                    {
                        _logWindow.Log($"처리되지 않은 명령: {command}");
                    }
                }
                else
                {
                    _logWindow.Log("메시지에 'command' 필드가 없습니다.");
                }
            }
            catch (Exception ex)
            {
                _logWindow.Log($"메시지 처리 중 오류 발생: {ex.Message}");
            }
        }
    }
} 