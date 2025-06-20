using System.Text.Json;
using overlay_gpt.Network.Models.Vue;
using overlay_gpt.Network.Models.Common;
using ProgramInfoCommon = overlay_gpt.Network.Models.Common.ProgramInfo;

namespace overlay_gpt.Services
{
    public class TextProcessingService
    {
        private readonly ChatDataManager _chatDataManager;

        public TextProcessingService()
        {
            _chatDataManager = ChatDataManager.Instance;
        }

        public DisplayText ProcessSelectedText(string selectedText, ProgramInfoCommon programInfo)
        {
            var now = DateTime.UtcNow;
            var timestamp = now.ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
            
            var chatData = new ChatData
            {
                GeneratedTimestamp = timestamp,
                CurrentProgram = programInfo,
                TargetProgram = null
            };

            _chatDataManager.AddChatData(chatData);

            var textData = new TextData { Type = "text_plain", Content = selectedText };

            return new DisplayText
            {
                GeneratedTimestamp = timestamp,
                CurrentProgram = programInfo,
                TargetProgram = null,
                Texts = new List<TextData> { textData }
            };
        }

        public string SerializeMessage(DisplayText message)
        {
            return JsonSerializer.Serialize(message, new JsonSerializerOptions 
            { 
                WriteIndented = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            });
        }
    }
} 