using System.Text.Json.Serialization;

namespace overlay_gpt.Network.Models.Common 
{
    public class ProgramInfo
    {

        [JsonPropertyName("context")]
        public string Context { get; set; } = string.Empty;

        [JsonPropertyName("fileId")]
        public ulong? FileId { get; set; }

        [JsonPropertyName("volumeId")]
        public uint? VolumeId { get; set; }

        [JsonPropertyName("fileType")]
        public string FileType { get; set; } = string.Empty;

        [JsonPropertyName("fileName")]
        public string FileName { get; set; } = string.Empty;

        [JsonPropertyName("filePath")]
        public string FilePath { get; set; } = string.Empty;

        [JsonPropertyName("position")]
        public string Position { get; set; } = string.Empty;
    } 
};

