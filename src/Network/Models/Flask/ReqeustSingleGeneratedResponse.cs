using System.Text.Json.Serialization;
using overlay_gpt.Network.Models.Common;

namespace overlay_gpt.Network.Models.Flask 
{
    public class RequestSingleGeneratedResponse
    {
        [JsonPropertyName("command")]
        public string Command { get; set; } = "request_single_generated_response";

        [JsonPropertyName("chat_id")]
        public int ChatId { get; set; }
        
        [JsonPropertyName("generated_timestamp")]
        public string GeneratedTimestamp { get; set; } = string.Empty;

        [JsonPropertyName("prompt")]
        public string Prompt { get; set; } = string.Empty;

        [JsonPropertyName("request_type")]
        public int RequestType { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; } = string.Empty;

        [JsonPropertyName("current_program")]
        public ProgramInfo CurrentProgram { get; set; } = new();

        [JsonPropertyName("target_program")]
        public ProgramInfo? TargetProgram { get; set; }
    } 
};

