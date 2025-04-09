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
                    LogWindow.Instance.Log($"TextPattern (선택): {text} (길이: {text.Length})");
                    
                    // 각 문자별로 스타일 정보 수집
                    var charStyles = new List<Dictionary<string, object>>();
                    for (int i = 0; i < text.Length; i++)
                    {
                        var charRange = selectedRange.Clone();
                        charRange.Move(TextUnit.Character, i);
                        charRange.ExpandToEnclosingUnit(TextUnit.Character);
                        charStyles.Add(GetStyleAttributes(charRange));
                    }
                    
                    // 문자별 스타일 정보를 포함하여 로깅
                    LogWindow.Instance.LogWithStylePerChar(text, charStyles);
                    return (text, styleAttributes);
                }
                else
                {
                    var fullText = textPattern.DocumentRange.GetText(-1);
                    LogWindow.Instance.Log($"TextPattern (전체): {fullText} (길이: {fullText.Length})");
                    
                    // 각 문자별로 스타일 정보 수집
                    var charStyles = new List<Dictionary<string, object>>();
                    for (int i = 0; i < fullText.Length; i++)
                    {
                        var charRange = textPattern.DocumentRange.Clone();
                        charRange.Move(TextUnit.Character, i);
                        charRange.ExpandToEnclosingUnit(TextUnit.Character);
                        charStyles.Add(GetStyleAttributes(charRange));
                    }
                    
                    // 문자별 스타일 정보를 포함하여 로깅
                    LogWindow.Instance.LogWithStylePerChar(fullText, charStyles);
                    return (fullText, styleAttributes);
                }
            }

            LogWindow.Instance.Log("No text pattern found");
            return (string.Empty, styleAttributes);
        }
    }
} 