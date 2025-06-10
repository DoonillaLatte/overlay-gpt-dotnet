using System.Text.Json.Serialization;
using Newtonsoft.Json;
using overlay_gpt.Network.Models.Common;

namespace overlay_gpt.Network.Models.Vue
{
    public class VueRequest
    {
        [JsonProperty("command")]
        [JsonPropertyName("command")]
        public string Command { get; set; } = "send_user_prompt";

        [JsonProperty("chat_id")]
        [JsonPropertyName("chat_id")]
        public int ChatId { get; set; }

        [JsonProperty("prompt")]
        [JsonPropertyName("prompt")]
        public string Prompt { get; set; } = string.Empty;

        [JsonProperty("current_program")]
        [JsonPropertyName("current_program")]
        public ProgramInfo CurrentProgram { get; set; } = new();

        [JsonProperty("target_program")]
        [JsonPropertyName("target_program")]
        public ProgramInfo? TargetProgram { get; set; }
    }
} 