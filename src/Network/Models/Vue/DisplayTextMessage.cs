using System.Collections.Generic;
using System.Text.Json.Serialization;
using overlay_gpt.Network.Models.Common;

namespace overlay_gpt.Network.Models.Vue
{
    public class DisplayText
    {
        [JsonPropertyName("command")]
        public string Command { get; set; } = "display_text";

        [JsonPropertyName("generated_timestamp")]
        public string GeneratedTimestamp { get; set; } = string.Empty;

        [JsonPropertyName("chat_id")]
        public int ChatId { get; set; } = -1;

        [JsonPropertyName("title")]
        public string Title { get; set; } = string.Empty;

        [JsonPropertyName("current_program")]
        public ProgramInfo CurrentProgram { get; set; } = new();

        [JsonPropertyName("target_program")]
        public ProgramInfo TargetProgram { get; set; } = new();

        [JsonPropertyName("texts")]
        public List<TextData> Texts { get; set; } = new();
    }
} 