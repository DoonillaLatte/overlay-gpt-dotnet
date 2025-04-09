using System;
using System.Collections.Generic;
using System.Windows.Automation;
using System.Windows.Automation.Text;

namespace overlay_gpt
{
    public class TextPatternContextReader : BaseContextReader
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

            if (element.TryGetCurrentPattern(TextPattern.Pattern, out object textPatternObj))
            {
                var textPattern = (TextPattern)textPatternObj;
                var selections = textPattern.GetSelection();

                if (selections != null && selections.Length > 0 && selections[0] != null && selections[0].GetText(-1).Length > 0)
                {
                    var selectedRange = selections[0];
                    string text = selectedRange.GetText(-1);
                    styleAttributes = GetStyleAttributes(selectedRange);
                    LogWindow.Instance.Log($"TextPattern (선택): {text} (길이: {text.Length})");
                    LogStyleAttributes(styleAttributes);
                    return (text, styleAttributes);
                }
                else
                {
                    var fullText = textPattern.DocumentRange.GetText(-1);
                    LogWindow.Instance.Log($"TextPattern (전체): {fullText} (길이: {fullText.Length})");
                    styleAttributes = GetStyleAttributes(textPattern.DocumentRange);
                    LogStyleAttributes(styleAttributes);
                    return (fullText, styleAttributes);
                }
            }

            LogWindow.Instance.Log("No text pattern found");
            return (string.Empty, styleAttributes);
        }
    }
} 