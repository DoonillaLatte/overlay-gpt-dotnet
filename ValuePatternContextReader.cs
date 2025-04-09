using System.Collections.Generic;
using System.Windows.Automation;

namespace overlay_gpt
{
    public class ValuePatternContextReader : BaseContextReader
    {
        public override (string SelectedText, Dictionary<string, object> StyleAttributes) GetSelectedTextWithStyle()
        {
            var styleAttributes = new Dictionary<string, object>();
            AutomationElement element = AutomationElement.FocusedElement;

            if (element == null)
            {
                LogWindow.Instance.Log("No focused element");
                return (string.Empty, styleAttributes);
            }

            if (element.TryGetCurrentPattern(ValuePattern.Pattern, out object valuePatternObj))
            {
                ValuePattern valuePattern = (ValuePattern)valuePatternObj;
                var value = valuePattern.Current.Value;
                LogWindow.Instance.Log($"ValuePattern: {value} (길이: {value.Length})");
                return (value, styleAttributes);
            }

            LogWindow.Instance.Log("No value pattern found");
            return (string.Empty, styleAttributes);
        }
    }
} 