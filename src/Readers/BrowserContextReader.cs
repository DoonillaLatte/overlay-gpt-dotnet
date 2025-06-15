using System;
using System.Collections.Generic;
using System.Windows.Automation;
using System.Runtime.InteropServices;
using System.Threading;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace overlay_gpt
{
    public class BrowserContextReader : BaseContextReader
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern bool GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        private const int MAX_RETRIES = 3;
        private const int RETRY_DELAY_MS = 300;

        public override (string SelectedText, Dictionary<string, object> StyleAttributes, string LineNumber)
            GetSelectedTextWithStyle(bool readAllContent = false)
        {
            var style = new Dictionary<string, object>();

            try
            {
                Console.WriteLine("BrowserContextReader 시작");
                
                // 현재 포커스된 창이 브라우저인지 확인
                if (!IsBrowserWindow())
                {
                    Console.WriteLine("현재 창이 브라우저가 아닙니다.");
                    return (string.Empty, style, string.Empty);
                }

                // 여러 방법으로 선택된 텍스트 가져오기 시도
                string selectedText = string.Empty;

                // 방법 1: Ctrl+C로 클립보드에 복사 후 가져오기
                selectedText = GetTextViaClipboard();
                
                if (string.IsNullOrEmpty(selectedText))
                {
                    // 방법 2: UI Automation을 사용하여 텍스트 가져오기
                    selectedText = GetTextViaUIAutomation();
                }

                Console.WriteLine($"선택된 텍스트 길이: {selectedText.Length}");
                return (selectedText, style, string.Empty);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"BrowserContextReader 오류: {ex.Message}");
                return (string.Empty, style, string.Empty);
            }
        }

        private bool IsBrowserWindow()
        {
            try
            {
                IntPtr foregroundWindow = GetForegroundWindow();
                if (foregroundWindow == IntPtr.Zero)
                    return false;

                var windowTitle = new StringBuilder(256);
                GetWindowText(foregroundWindow, windowTitle, windowTitle.Capacity);
                string title = windowTitle.ToString().ToLower();

                // 주요 브라우저 창 제목 패턴 확인
                return title.Contains("chrome") || 
                       title.Contains("firefox") || 
                       title.Contains("edge") || 
                       title.Contains("safari") || 
                       title.Contains("opera") || 
                       title.Contains("브라우저") ||
                       title.Contains("internet explorer") ||
                       title.Contains("mozilla") ||
                       title.Contains("webkit");
            }
            catch
            {
                return false;
            }
        }

        private string GetTextViaClipboard()
        {
            try
            {
                Console.WriteLine("클립보드를 통한 텍스트 가져오기 시도");
                
                // 현재 클립보드 내용 백업
                string originalText = string.Empty;
                string originalHtml = string.Empty;
                try
                {
                    if (Clipboard.ContainsText())
                        originalText = Clipboard.GetText();
                    if (Clipboard.ContainsText(TextDataFormat.Html))
                        originalHtml = Clipboard.GetText(TextDataFormat.Html);
                }
                catch { }

                // 클립보드 비우기
                try
                {
                    Clipboard.Clear();
                    Thread.Sleep(100);
                }
                catch { }

                // Ctrl+C 전송하여 선택된 텍스트 복사
                SendKeys.SendWait("^c");
                Thread.Sleep(500); // HTML 형식 데이터 처리를 위해 대기 시간 증가

                string selectedText = string.Empty;
                
                for (int i = 0; i < MAX_RETRIES; i++)
                {
                    try
                    {
                        // HTML 형식이 있는지 먼저 확인
                        if (Clipboard.ContainsText(TextDataFormat.Html))
                        {
                            string htmlContent = Clipboard.GetText(TextDataFormat.Html);
                            Console.WriteLine($"HTML 클립보드 데이터 길이: {htmlContent.Length}");
                            
                            // HTML 클립보드 데이터에서 실제 HTML 부분 추출
                            string extractedHtml = ExtractHtmlFromClipboard(htmlContent);
                            
                            if (!string.IsNullOrEmpty(extractedHtml))
                            {
                                // HTML을 그대로 반환 (파싱하지 않음)
                                selectedText = extractedHtml;
                                if (!string.IsNullOrEmpty(selectedText))
                                    break;
                            }
                        }
                        
                        // HTML이 없거나 처리 실패 시 일반 텍스트 사용
                        if (string.IsNullOrEmpty(selectedText) && Clipboard.ContainsText())
                        {
                            selectedText = Clipboard.GetText();
                            if (!string.IsNullOrEmpty(selectedText))
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"클립보드 읽기 재시도 {i + 1}: {ex.Message}");
                    }
                    
                    Thread.Sleep(RETRY_DELAY_MS);
                }

                // 원래 클립보드 내용 복원
                try
                {
                    if (!string.IsNullOrEmpty(originalHtml))
                    {
                        Clipboard.SetText(originalHtml, TextDataFormat.Html);
                    }
                    else if (!string.IsNullOrEmpty(originalText))
                    {
                        Clipboard.SetText(originalText);
                    }
                }
                catch { }

                Console.WriteLine($"클립보드에서 가져온 텍스트 길이: {selectedText.Length}");
                return selectedText;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"클립보드 방식 실패: {ex.Message}");
                return string.Empty;
            }
        }

        private string ExtractHtmlFromClipboard(string clipboardHtml)
        {
            try
            {
                // Windows 클립보드 HTML 형식에서 실제 HTML 부분 추출
                // 클립보드 HTML은 "StartHTML:xxxxx\r\nEndHTML:xxxxx\r\n..." 형식
                
                var startHtmlMatch = Regex.Match(clipboardHtml, @"StartHTML:(\d+)");
                var endHtmlMatch = Regex.Match(clipboardHtml, @"EndHTML:(\d+)");
                
                if (startHtmlMatch.Success && endHtmlMatch.Success)
                {
                    int startIndex = int.Parse(startHtmlMatch.Groups[1].Value);
                    int endIndex = int.Parse(endHtmlMatch.Groups[1].Value);
                    
                    if (startIndex < clipboardHtml.Length && endIndex <= clipboardHtml.Length && endIndex > startIndex)
                    {
                        return clipboardHtml.Substring(startIndex, endIndex - startIndex);
                    }
                }
                
                // 매칭이 실패하면 HTML 태그가 있는 부분을 찾아서 반환
                var htmlTagMatch = Regex.Match(clipboardHtml, @"<html.*?</html>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                if (htmlTagMatch.Success)
                {
                    return htmlTagMatch.Value;
                }
                
                // 그래도 없으면 body 태그만 찾기
                var bodyTagMatch = Regex.Match(clipboardHtml, @"<body.*?</body>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                if (bodyTagMatch.Success)
                {
                    return bodyTagMatch.Value;
                }
                
                return clipboardHtml;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"HTML 추출 실패: {ex.Message}");
                return clipboardHtml;
            }
        }



        private string GetTextViaUIAutomation()
        {
            try
            {
                Console.WriteLine("UI Automation을 통한 텍스트 가져오기 시도");
                
                var element = AutomationElement.FocusedElement;
                if (element == null)
                    return string.Empty;

                // TextPattern을 지원하는지 확인
                if (element.TryGetCurrentPattern(TextPattern.Pattern, out object textPatternObj))
                {
                    var textPattern = textPatternObj as TextPattern;
                    var selection = textPattern.GetSelection();
                    
                    if (selection != null && selection.Length > 0)
                    {
                        var selectedText = selection[0].GetText(-1);
                        Console.WriteLine($"UI Automation에서 가져온 텍스트 길이: {selectedText.Length}");
                        return selectedText;
                    }
                }

                return string.Empty;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"UI Automation 방식 실패: {ex.Message}");
                return string.Empty;
            }
        }

        public override (ulong? FileId, uint? VolumeId, string FileType, string FileName, string FilePath) GetFileInfo()
        {
            try
            {
                IntPtr foregroundWindow = GetForegroundWindow();
                var windowTitle = new StringBuilder(256);
                GetWindowText(foregroundWindow, windowTitle, windowTitle.Capacity);
                
                return (null, null, "Browser", "웹 페이지", windowTitle.ToString());
            }
            catch
            {
                return (null, null, "Browser", "웹 페이지", "Unknown");
            }
        }

        private string ProcessImageUrls(string htmlContent)
        {
            try
            {
                Console.WriteLine("이미지 URL 처리 시작");
                
                // 현재 브라우저의 URL을 가져와서 base URL로 사용
                string baseUrl = GetCurrentBrowserUrl();
                
                // img 태그의 src 속성을 찾아서 처리
                string processedHtml = Regex.Replace(htmlContent, 
                    @"<img([^>]*?)src\s*=\s*[""']([^""']*?)[""']([^>]*?)>", 
                    match => ProcessImageTag(match, baseUrl), 
                    RegexOptions.IgnoreCase);

                Console.WriteLine($"이미지 처리 완료: {processedHtml.Length} 문자");
                return processedHtml;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"이미지 URL 처리 실패: {ex.Message}");
                return htmlContent; // 처리 실패 시 원본 반환
            }
        }

        private string ProcessImageTag(Match match, string baseUrl)
        {
            try
            {
                string beforeSrc = match.Groups[1].Value;
                string srcValue = match.Groups[2].Value;
                string afterSrc = match.Groups[3].Value;

                Console.WriteLine($"원본 이미지 URL: {srcValue}");

                // 이미지 URL 처리
                string processedSrc = ProcessSingleImageUrl(srcValue, baseUrl);
                
                Console.WriteLine($"처리된 이미지 URL: {processedSrc}");

                // 이미지 태그 재구성
                return $"<img{beforeSrc}src=\"{processedSrc}\"{afterSrc}>";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"이미지 태그 처리 실패: {ex.Message}");
                return match.Value; // 처리 실패 시 원본 반환
            }
        }

        private string ProcessSingleImageUrl(string srcValue, string baseUrl)
        {
            // 이미 완전한 URL인 경우
            if (srcValue.StartsWith("http://") || srcValue.StartsWith("https://"))
            {
                return srcValue;
            }
            
            // data: URL (base64)인 경우
            if (srcValue.StartsWith("data:"))
            {
                return srcValue;
            }
            
            // 상대 경로인 경우 절대 URL로 변환
            if (!string.IsNullOrEmpty(baseUrl))
            {
                try
                {
                    var baseUri = new Uri(baseUrl);
                    var absoluteUri = new Uri(baseUri, srcValue);
                    return absoluteUri.ToString();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"URL 변환 실패: {ex.Message}");
                }
            }
            
            return srcValue; // 변환 실패 시 원본 반환
        }

        private string GetCurrentBrowserUrl()
        {
            try
            {
                IntPtr foregroundWindow = GetForegroundWindow();
                if (foregroundWindow == IntPtr.Zero)
                    return string.Empty;

                // UI Automation을 사용하여 주소창 URL 가져오기
                var element = AutomationElement.FromHandle(foregroundWindow);
                if (element == null)
                    return string.Empty;

                // Chrome의 경우 주소창에서 URL 추출 시도
                var addressBar = element.FindFirst(TreeScope.Descendants,
                    new AndCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit),
                        new PropertyCondition(AutomationElement.NameProperty, "주소 및 검색 창")
                    ));

                if (addressBar == null)
                {
                    // 다른 언어/브라우저의 경우 - 영어로 시도
                    addressBar = element.FindFirst(TreeScope.Descendants,
                        new AndCondition(
                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit),
                            new PropertyCondition(AutomationElement.NameProperty, "Address and search bar")
                        ));
                }

                if (addressBar == null)
                {
                    // 일반적인 Edit 컨트롤 중에서 URL 형태의 값을 가진 것 찾기
                    var editControls = element.FindAll(TreeScope.Descendants,
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit));
                        
                    foreach (AutomationElement edit in editControls)
                    {
                        try
                        {
                            var valuePattern = edit.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                            if (valuePattern != null)
                            {
                                string value = valuePattern.Current.Value;
                                if (!string.IsNullOrEmpty(value) && 
                                    (value.StartsWith("http://") || value.StartsWith("https://")))
                                {
                                    addressBar = edit;
                                    break;
                                }
                            }
                        }
                        catch { continue; }
                    }
                }

                if (addressBar != null)
                {
                    var valuePattern = addressBar.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                    if (valuePattern != null)
                    {
                        string url = valuePattern.Current.Value;
                        if (!string.IsNullOrEmpty(url) && (url.StartsWith("http://") || url.StartsWith("https://")))
                        {
                            Console.WriteLine($"브라우저 URL 감지: {url}");
                            return url;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"브라우저 URL 가져오기 실패: {ex.Message}");
            }

            return string.Empty;
        }
    }
} 