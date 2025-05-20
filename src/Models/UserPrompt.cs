using System;
using System.Text.Json.Serialization;

namespace overlay_gpt.Models
{
    public class UserPrompt
    {
        [JsonPropertyName("chat_id")]
        public int ChatId { get; set; }

        [JsonPropertyName("prompt")]
        public string Prompt { get; set; } = string.Empty;

        [JsonPropertyName("target_program")]
        public string? TargetProgram { get; set; }
    }
} 