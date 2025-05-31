using System.Collections.Generic;
using System.Windows.Automation;

namespace overlay_gpt
{
    public class ValuePatternContextReader : BaseContextReader
    {
        public override (string SelectedText, Dictionary<string, object> StyleAttributes, string LineNumber) GetSelectedTextWithStyle()
        {
            var styleAttributes = new Dictionary<string, object>();
            AutomationElement element = AutomationElement.FocusedElement;

            if (element == null)
            {
                LogWindow.Instance.Log("No focused element");
                return (string.Empty, styleAttributes, string.Empty);
            }

            if (element.TryGetCurrentPattern(ValuePattern.Pattern, out object valuePatternObj))
            {
                var valuePattern = (ValuePattern)valuePatternObj;
                string text = valuePattern.Current.Value;
                return (text, styleAttributes, string.Empty);
            }

            return (string.Empty, styleAttributes, string.Empty);
        }
    }
} 