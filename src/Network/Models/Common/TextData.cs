using System.Text.Json.Serialization;
using Newtonsoft.Json;

namespace overlay_gpt.Network.Models.Common
{
    public class TextData
    {
        [JsonProperty("type")]
        [JsonPropertyName("type")]
        public string Type { get; set; } = "text_plain";
        [JsonProperty("content")]
        [JsonPropertyName("content")]
        public string Content { get; set; } = string.Empty;
    }
}