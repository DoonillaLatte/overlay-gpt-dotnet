using System;
using System.Threading.Tasks;
using System.Text.Json;
using overlay_gpt.Services;

namespace overlay_gpt.Scripts
{
    public static class WebSocketMessageSender
    {
        /// <summary>
        /// 채팅 메시지를 전송합니다.
        /// </summary>
        /// <param name="message">전송할 메시지</param>
        public static async Task SendChatMessageAsync(string message)
        {
            try
            {
                var parameters = new { message = message };
                await overlay_gpt.Services.WebSocketManager.Instance.SendMessageAsync("chat", parameters);
                LogWindow.Instance.Log($"채팅 메시지 전송: {message}");
            }
            catch (Exception ex)
            {
                LogWindow.Instance.Log($"채팅 메시지 전송 실패: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 명령을 전송합니다.
        /// </summary>
        /// <param name="command">명령어 이름</param>
        /// <param name="parameters">명령어 파라미터 (JSON 객체)</param>
        public static async Task SendCommandAsync(string command, object parameters)
        {
            try
            {
                var message = new
                {
                    command = command,
                    parameters = parameters
                };

                await overlay_gpt.Services.WebSocketManager.Instance.SendMessageAsync(command, message);
                LogWindow.Instance.Log($"명령 전송: {command} - {JsonSerializer.Serialize(parameters)}");
            }
            catch (Exception ex)
            {
                LogWindow.Instance.Log($"명령 전송 실패: {ex.Message}");
                throw;
            }
        }
    }
} 