using System.Text.Json.Serialization;
using Newtonsoft.Json;

namespace overlay_gpt.Network.Models.Common
{
    public class ChatData
    {
        [JsonProperty("chat_id")]
        [JsonPropertyName("chat_id")]
        public int ChatId { get; set; } = -1;
        [JsonProperty("title")]
        [JsonPropertyName("title")]
        public string Title { get; set; } = string.Empty;
        [JsonProperty("generated_timestamp")]
        [JsonPropertyName("generated_timestamp")]
        public string GeneratedTimestamp { get; set; } = string.Empty;
        [JsonProperty("current_program")]
        [JsonPropertyName("current_program")]
        public ProgramInfo CurrentProgram { get; set; } = new();
        [JsonProperty("target_program")]
        [JsonPropertyName("target_program")]
        public ProgramInfo? TargetProgram { get; set; }
        [JsonProperty("texts")]
        [JsonPropertyName("texts")]
        public List<TextData> Texts { get; set; } = new();
        
        // Vue 표시용 정규화된 HTML (Vue에만 전송)
        public string VueDisplayContext { get; set; } = string.Empty;
        
        // dotnet 적용용 원본 HTML (실제 적용에 사용)
        public string DotnetApplyContext { get; set; } = string.Empty;
    }
}