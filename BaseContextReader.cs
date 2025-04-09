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
            LogWindow.Instance.Log($"Style - FontName: {styleAttributes["FontName"]}, FontSize: {styleAttributes["FontSize"]}, " +
                               $"FontWeight: {styleAttributes["FontWeight"]}, ForegroundColor: {styleAttributes["ForegroundColor"]}, " +
                               $"BackgroundColor: {styleAttributes["BackgroundColor"]}, UnderlineStyle: {styleAttributes["UnderlineStyle"]}");
        }

        public abstract (string SelectedText, Dictionary<string, object> StyleAttributes) GetSelectedTextWithStyle();
    }
} 