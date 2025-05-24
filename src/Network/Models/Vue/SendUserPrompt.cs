using System.Text.Json.Serialization;
using overlay_gpt.Network.Models.Common;

namespace overlay_gpt.Network.Models.Vue
{
    public class VueRequest
    {
        [JsonPropertyName("command")]
        public string Command { get; set; } = "send_user_prompt";

        [JsonPropertyName("chat_id")]
        public int ChatId { get; set; }

        [JsonPropertyName("prompt")]
        public string Prompt { get; set; } = string.Empty;

        [JsonPropertyName("current_program")]
        public ProgramInfo CurrentProgram { get; set; } = new();

        [JsonPropertyName("target_program")]
        public ProgramInfo? TargetProgram { get; set; }
    }
} 