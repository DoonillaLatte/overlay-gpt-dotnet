using System.Collections.Generic;

namespace overlay_gpt
{
    public interface IContextReader
    {
        (string SelectedText, Dictionary<string, object> StyleAttributes) GetSelectedTextWithStyle();
    }
} 