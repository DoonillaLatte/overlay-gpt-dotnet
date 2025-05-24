using System.Collections.Generic;
using System.Text.Json.Serialization;
using overlay_gpt.Network.Models.Common;

namespace overlay_gpt.Network.Models.Vue
{
    public class DisplayText
    {
        [JsonPropertyName("command")]
        public string Command { get; set; } = "display_text";

        [JsonPropertyName("chat_id")]
        public int ChatId { get; set; } = 1;

        [JsonPropertyName("current_program")]
        public ProgramInfo CurrentProgram { get; set; } = new();

        [JsonPropertyName("target_program")]
        public ProgramInfo TargetProgram { get; set; } = new();

        [JsonPropertyName("texts")]
        public List<TextInfo> Texts { get; set; } = new();
    }

    public class TextInfo
    {
        [JsonPropertyName("type")]
        public string Type { get; set; } = "text_plain";

        [JsonPropertyName("content")]
        public object Content { get; set; } = string.Empty;
    }
} 