using System.Text.Json.Serialization;

namespace overlay_gpt.Network.Models.Common 
{
    public class ProgramInfo
    {
        [JsonPropertyName("id")]
        public int Id { get; set; } = -1;

        [JsonPropertyName("type")]
        public string Type { get; set; } = string.Empty;

        [JsonPropertyName("context")]
        public string Context { get; set; } = string.Empty;
    } 
};

