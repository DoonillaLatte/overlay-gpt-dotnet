using System.Collections.Generic;
using System.Windows.Automation;
using System.Windows.Automation.Text;

namespace overlay_gpt
{
    public abstract class BaseContextReader : IContextReader
    {
        protected Dictionary<string, object> GetStyleAttributes(TextPatternRange range)
        {
            var styleAttributes = new Dictionary<string, object>();
            
            styleAttributes["FontName"] = range.GetAttributeValue(TextPattern.FontNameAttribute);
            styleAttributes["FontSize"] = range.GetAttributeValue(TextPattern.FontSizeAttribute);
            styleAttributes["FontWeight"] = range.GetAttributeValue(TextPattern.FontWeightAttribute);
            styleAttributes["ForegroundColor"] = range.GetAttributeValue(TextPattern.ForegroundColorAttribute);
            styleAttributes["BackgroundColor"] = range.GetAttributeValue(TextPattern.BackgroundColorAttribute);
            styleAttributes["UnderlineStyle"] = range.GetAttributeValue(TextPattern.UnderlineStyleAttribute);

            return styleAttributes;
        }

        protected void LogStyleAttributes(Dictionary<string, object> styleAttributes)
        {
            string styleInfo = $"Style - FontName: {styleAttributes["FontName"]}, FontSize: {styleAttributes["FontSize"]}, " +
                             $"FontWeight: {styleAttributes["FontWeight"]}, ForegroundColor: {styleAttributes["ForegroundColor"]}, " +
                             $"BackgroundColor: {styleAttributes["BackgroundColor"]}, UnderlineStyle: {styleAttributes["UnderlineStyle"]}";
            
            LogWindow.Instance.LogWithStyle(styleInfo, styleAttributes);
        }

        public abstract (string SelectedText, Dictionary<string, object> StyleAttributes) GetSelectedTextWithStyle();

        // 파일 정보를 가져오는 메서드
        public virtual (ulong? FileId, uint? VolumeId, string FileType, string FileName, string FilePath) GetFileInfo()
        {
            return (null, null, string.Empty, string.Empty, string.Empty);
        }
    }
} 