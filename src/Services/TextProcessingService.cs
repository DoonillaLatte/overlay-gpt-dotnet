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
            var chatData = new ChatData
            {
                GeneratedTimestamp = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                CurrentProgram = programInfo,
                TargetProgram = null
            };

            _chatDataManager.AddChatData(chatData);

            return new DisplayText
            {
                CurrentProgram = programInfo,
                TargetProgram = null,
                Texts = new List<TextInfo> {}
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