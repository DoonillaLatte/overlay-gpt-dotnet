using System.Windows.Automation;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using System.Diagnostics;
using System;

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
                
            // Word 리더 추가
            try
            {   
                Console.WriteLine("WordContextReader 생성 시도");
                var wordReader = new WordContextReader();
                var (text, _, _) = wordReader.GetSelectedTextWithStyle();
                if (!string.IsNullOrEmpty(text))
                {
                    Console.WriteLine("WordContextReader 생성 성공");
                    return wordReader;
                }
                throw new InvalidOperationException("No text selected in Word");
            }
            catch(Exception e)
            {
                Console.WriteLine($"Word 관련 오류 발생: {e.Message}");
            }

            // Excel 리더 추가
            try
            {   
                Console.WriteLine("ExcelContextReader 생성 시도");
                var excelReader = new ExcelContextReader();
                var (text, _, _) = excelReader.GetSelectedTextWithStyle();
                if (!string.IsNullOrEmpty(text))
                {
                    Console.WriteLine("ExcelContextReader 생성 성공");
                    return excelReader;
                }
                throw new InvalidOperationException("No text selected in Excel");
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