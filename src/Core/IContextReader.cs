using System.Collections.Generic;

namespace overlay_gpt
{
    public interface IContextReader
    {
        (string SelectedText, Dictionary<string, object> StyleAttributes, string LineNumber) GetSelectedTextWithStyle(bool readAllContent = false);
        (ulong? FileId, uint? VolumeId, string FileType, string FileName, string FilePath) GetFileInfo();
    }
} 