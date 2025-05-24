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

        public DisplayText ProcessSelectedText(string selectedText, string fileName = "", string programType = "")
        {
            var chatData = new ChatData
            {
                ChatId = -1,
                GeneratedTimestamp = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                CurrentProgram = new ProgramInfoCommon
                {
                    Id = -1,
                    Type = programType,
                    Context = selectedText
                },
                TargetProgram = null
            };

            _chatDataManager.AddChatData(chatData);

            return new DisplayText
            {
                ChatId = chatData.ChatId,
                CurrentProgram = new ProgramInfo
                {
                    Id = -1,
                    Type = programType,
                    Context = selectedText
                },
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