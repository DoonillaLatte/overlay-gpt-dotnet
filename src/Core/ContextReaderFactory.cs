using System.Windows.Automation;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using System.Diagnostics;

namespace overlay_gpt
{
    public static class ContextReaderFactory
    {
        public static IContextReader CreateReader(AutomationElement element)
        {
            if (element == null)
            {
                Console.WriteLine("element가 null입니다.");
                return new TextPatternContextReader();
            }
                
            // Excel 리더 추가
            try
            {   
                Console.WriteLine("ExcelContextReader 생성 시도");
                var reader = new ExcelContextReader();
                Console.WriteLine("ExcelContextReader 생성 성공");
                return reader;
            }
            catch(Exception e)
            {
                Console.WriteLine($"Excel 관련 오류 발생: {e.Message}");
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