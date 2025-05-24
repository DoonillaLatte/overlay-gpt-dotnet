using System.Text.Json.Serialization;

namespace overlay_gpt.Network.Models.Common
{
    public class ChatData
    {
        [JsonPropertyName("chat_id")]
        public int ChatId { get; set; } = -1;
        [JsonPropertyName("generated_timestamp")]
        public string GeneratedTimestamp { get; set; } = string.Empty;
        [JsonPropertyName("current_program")]
        public ProgramInfo CurrentProgram { get; set; } = new();
        [JsonPropertyName("target_program")]
        public ProgramInfo? TargetProgram { get; set; }
    }
}