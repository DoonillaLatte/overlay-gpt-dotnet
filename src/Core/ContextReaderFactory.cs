using System.Windows.Automation;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;

namespace overlay_gpt
{
    public static class ContextReaderFactory
    {
        public static IContextReader CreateReader(AutomationElement element)
        {
            if (element == null)
                return new TextPatternContextReader();
                
            // Excel 리더 추가
            try
            {   
                return new ExcelContextReader();
            }
            catch(Exception e)
            {
                Console.WriteLine("Excel 관련 오류 발생");
                Console.WriteLine(e.Message);
            }
                
            // TextBox나 ValueBox일 때 포커스 여부 확인
            if (element.TryGetCurrentPattern(TextPattern.Pattern, out _) || 
                element.TryGetCurrentPattern(ValuePattern.Pattern, out _))
            {
                // 현재 포커스된 요소와 비교
                var focusedElement = AutomationElement.FocusedElement;
                if (focusedElement != null && focusedElement.Equals(element))
                {
                    if (element.TryGetCurrentPattern(TextPattern.Pattern, out _))
                        return new TextPatternContextReader();
                    
                    if (element.TryGetCurrentPattern(ValuePattern.Pattern, out _))
                        return new ValuePatternContextReader();
                }
                return new ClipboardContextReader();
            }

            // 클립보드 리더 추가
            return new ClipboardContextReader();
        }
    }
} 