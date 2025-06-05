using System;
using System.Collections.Generic;
using System.Windows.Automation;
using System.Windows.Automation.Text;

namespace overlay_gpt
{
    public class TextPatternContextReader : BaseContextReader
    {
        public override (string SelectedText, Dictionary<string, object> StyleAttributes, string LineNumber) GetSelectedTextWithStyle(bool readAllContent = false)
        {
            var styleAttributes = new Dictionary<string, object>();
            AutomationElement element = AutomationElement.FocusedElement;

            if (element == null)
            {
                LogWindow.Instance.Log("No focused element");
                return (string.Empty, styleAttributes, string.Empty);
            }

            if (element.TryGetCurrentPattern(TextPattern.Pattern, out object textPatternObj))
            {
                var textPattern = (TextPattern)textPatternObj;
                var selections = textPattern.GetSelection();
                
                if (selections.Length > 0)
                {
                    var range = selections[0];
                    string text = range.GetText(-1);
                    styleAttributes = GetStyleAttributes(range);
                    
                    // 줄 번호 정보 가져오기 (TextPattern에서는 줄 번호를 직접 가져올 수 없으므로 빈 문자열 반환)
                    string lineNumber = string.Empty;
                    
                    return (text, styleAttributes, lineNumber);
                }
            }

            return (string.Empty, styleAttributes, string.Empty);
        }
    }
} 