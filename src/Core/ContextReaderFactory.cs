using System.Windows.Automation;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using System.Diagnostics;
using System;
using Microsoft.Extensions.Logging;
using System.IO;

namespace overlay_gpt
{
    public static class ContextReaderFactory
    {
        private static readonly ILoggerFactory _loggerFactory = LoggerFactory.Create(builder =>
        {
            builder.AddConsole();
        });

        public static IContextReader CreateReader(AutomationElement element, bool isTargetProg = false, string filePath = "")
        {
            if (element == null)
            {
                Console.WriteLine("element가 null입니다.");
                return new TextPatternContextReader();
            }

            // 파일 경로가 있는 경우 확장자 확인
            if (!string.IsNullOrEmpty(filePath))
            {
                string extension = Path.GetExtension(filePath).ToLower();
                
                // 한글 리더 추가
                if (extension == ".hwp")
                {
                    try
                    {   
                        Console.WriteLine("HwpContextReader 생성 시도");
                        var logger = _loggerFactory.CreateLogger<HwpContextReader>();
                        var hwpReader = new HwpContextReader(logger);
                        var (text, _, _) = hwpReader.GetSelectedTextWithStyle();
                        if (!string.IsNullOrEmpty(text))
                        {
                            Console.WriteLine("HwpContextReader 생성 성공");
                            return hwpReader;
                        }
                        throw new InvalidOperationException("No text selected in Hwp");
                    }
                    catch(Exception e)
                    {
                        Console.WriteLine($"한글 관련 오류 발생: {e.Message}");
                    }
                }

                // Word 리더 추가
                if (extension == ".docx" || extension == ".doc")
                {
                    try
                    {   
                        Console.WriteLine("WordContextReader 생성 시도");
                        var wordReader = new WordContextReader(isTargetProg, filePath);
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
                }

                // Excel 리더 추가
                if (extension == ".xlsx" || extension == ".xls")
                {
                    try
                    {   
                        Console.WriteLine("ExcelContextReader 생성 시도");
                        var excelReader = new ExcelContextReader(isTargetProg, filePath);
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
                }

                // PowerPoint 리더 추가
                if (extension == ".pptx" || extension == ".ppt")
                {
                    try
                    {   
                        Console.WriteLine("PPTContextReader 생성 시도");
                        var pptReader = new PPTContextReader(isTargetProg, filePath);
                        var (text, _, _) = pptReader.GetSelectedTextWithStyle();
                        if (!string.IsNullOrEmpty(text))
                        {
                            Console.WriteLine("PPTContextReader 생성 성공");
                            return pptReader;
                        }
                        throw new InvalidOperationException("No text selected in PowerPoint");
                    }
                    catch(Exception e)
                    {
                        Console.WriteLine($"PowerPoint 관련 오류 발생: {e.Message}");
                    }
                }
            }
            else
            {
                // 한글 리더 추가
                try
                {   
                    Console.WriteLine("HwpContextReader 생성 시도");
                    var logger = _loggerFactory.CreateLogger<HwpContextReader>();
                    var hwpReader = new HwpContextReader(logger);
                    var (text, _, _) = hwpReader.GetSelectedTextWithStyle();
                    if (!string.IsNullOrEmpty(text))
                    {
                        Console.WriteLine("HwpContextReader 생성 성공");
                        return hwpReader;
                    }
                    throw new InvalidOperationException("No text selected in Hwp");
                }
                catch(Exception e)
                {
                    Console.WriteLine($"한글 관련 오류 발생: {e.Message}");
                }

                // Word 리더 추가
                try
                {   
                    Console.WriteLine("WordContextReader 생성 시도");
                    var wordReader = new WordContextReader(isTargetProg, filePath);
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
                    var excelReader = new ExcelContextReader(isTargetProg, filePath);
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

                // PowerPoint 리더 추가
                try
                {   
                    Console.WriteLine("PPTContextReader 생성 시도");
                    var pptReader = new PPTContextReader();
                    var (text, _, _) = pptReader.GetSelectedTextWithStyle();
                    if (!string.IsNullOrEmpty(text))
                    {
                        Console.WriteLine("PPTContextReader 생성 성공");
                        return pptReader;
                    }
                    throw new InvalidOperationException("No text selected in PowerPoint");
                }
                catch(Exception e)
                {
                    Console.WriteLine($"PowerPoint 관련 오류 발생: {e.Message}");
                }
            }
                
            // TextBox나 ValueBox일 때 포커스 여부 확인
            /*if (element.TryGetCurrentPattern(TextPattern.Pattern, out _) || 
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
            }*/

            // 클립보드 리더 추가
            return new ClipboardContextReader();
        }
    }
} 