using System;
using System.Collections.Generic;
using System.Windows.Automation;
using Microsoft.Office.Interop.Word;
using WordFont = Microsoft.Office.Interop.Word.Font;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.IO;
using Forms = System.Windows.Forms;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using System.Linq;
using System.Windows.Forms;
using WordApp = Microsoft.Office.Interop.Word.Application;
using HtmlDoc = HtmlAgilityPack.HtmlDocument;

namespace overlay_gpt
{
    public class WordContextReader : BaseContextReader
    {
        private WordApp? _wordApp;
        private Document? _document;
        private string? _filePath;
        private bool _isTargetProg;

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("oleaut32.dll")]
        private static extern int GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        [DllImport("ole32.dll")]
        private static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr CreateFile(
            string lpFileName,
            uint dwDesiredAccess,
            uint dwShareMode,
            IntPtr lpSecurityAttributes,
            uint dwCreationDisposition,
            uint dwFlagsAndAttributes,
            IntPtr hTemplateFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool GetFileInformationByHandle(
            IntPtr hFile,
            out BY_HANDLE_FILE_INFORMATION lpFileInformation);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr hObject);

        [StructLayout(LayoutKind.Sequential)]
        private struct BY_HANDLE_FILE_INFORMATION
        {
            public uint dwFileAttributes;
            public FILETIME ftCreationTime;
            public FILETIME ftLastAccessTime;
            public FILETIME ftLastWriteTime;
            public uint dwVolumeSerialNumber;
            public uint nFileSizeHigh;
            public uint nFileSizeLow;
            public uint nNumberOfLinks;
            public uint nFileIndexHigh;
            public uint nFileIndexLow;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct FILETIME
        {
            public uint dwLowDateTime;
            public uint dwHighDateTime;
        }

        private const uint GENERIC_READ = 0x80000000;
        private const uint FILE_SHARE_READ = 0x00000001;
        private const uint FILE_SHARE_WRITE = 0x00000002;
        private const uint OPEN_EXISTING = 3;

        public WordContextReader(bool isTargetProg = false, string filePath = "")
        {
            Console.WriteLine($"WordContextReader 생성 시도 - isTargetProg: {isTargetProg}");
            _isTargetProg = isTargetProg;
            _filePath = filePath;
        }

        private static object GetActiveObject(string progID)
        {
            Guid clsid;
            CLSIDFromProgID(progID, out clsid);
            object obj;
            GetActiveObject(ref clsid, IntPtr.Zero, out obj);
            return obj;
        }

        private (ulong FileId, uint VolumeId)? GetFileId(string filePath)
        {
            try
            {
                Console.WriteLine($"GetFileId 호출 - 파일 경로: {filePath}");
                
                if (!File.Exists(filePath))
                {
                    Console.WriteLine("파일이 존재하지 않습니다.");
                    return null;
                }

                IntPtr handle = CreateFile(
                    filePath,
                    GENERIC_READ,
                    FILE_SHARE_READ | FILE_SHARE_WRITE,
                    IntPtr.Zero,
                    OPEN_EXISTING,
                    0,
                    IntPtr.Zero);

                if (handle.ToInt64() == -1)
                {
                    Console.WriteLine($"CreateFile 실패 - 에러 코드: {Marshal.GetLastWin32Error()}");
                    return null;
                }

                try
                {
                    BY_HANDLE_FILE_INFORMATION fileInfo;
                    if (GetFileInformationByHandle(handle, out fileInfo))
                    {
                        ulong fileId = ((ulong)fileInfo.nFileIndexHigh << 32) | fileInfo.nFileIndexLow;
                        Console.WriteLine($"파일 ID 정보 가져오기 성공:");
                        Console.WriteLine($"- FileId: {fileId}");
                        Console.WriteLine($"- VolumeId: {fileInfo.dwVolumeSerialNumber}");
                        return (fileId, fileInfo.dwVolumeSerialNumber);
                    }
                    else
                    {
                        Console.WriteLine($"GetFileInformationByHandle 실패 - 에러 코드: {Marshal.GetLastWin32Error()}");
                    }
                }
                finally
                {
                    CloseHandle(handle);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"파일 ID 가져오기 오류: {ex.Message}");
                Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
            }
            return null;
        }

        private bool IsWordProcessActive()
        {
            try
            {
                IntPtr foregroundWindow = GetForegroundWindow();
                if (foregroundWindow == IntPtr.Zero)
                {
                    Console.WriteLine("포커스된 창을 찾을 수 없습니다.");
                    return false;
                }

                uint processId;
                GetWindowThreadProcessId(foregroundWindow, out processId);

                Process foregroundProcess = Process.GetProcessById((int)processId);
                Console.WriteLine($"현재 포커스된 프로세스: {foregroundProcess.ProcessName} (PID: {processId})");

                // Word 프로세스 이름 확인 (WINWORD.EXE)
                return foregroundProcess.ProcessName.Equals("WINWORD", StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"프로세스 확인 중 오류 발생: {ex.Message}");
                return false;
            }
        }

        public override (string SelectedText, Dictionary<string, object> StyleAttributes, string LineNumber) GetSelectedTextWithStyle(bool includeStyle = true)
        {
            try
            {
                Console.WriteLine("Word 데이터 읽기 시작...");

                // isTargetProg가 false일 때만 현재 포커스된 프로세스가 Word인지 확인
                if (!_isTargetProg && !IsWordProcessActive())
                {
                    Console.WriteLine("현재 포커스된 프로세스가 Word가 아닙니다.");
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);
                }

                // Word COM 객체 생성 시도
                try
                {
                    Console.WriteLine("기존 Word 애플리케이션 찾기 시도...");
                    _wordApp = (WordApp)GetActiveObject("Word.Application");
                    
                    if (_wordApp != null)
                    {
                        Console.WriteLine("기존 Word 애플리케이션 찾음");
                        Console.WriteLine($"Word 버전: {_wordApp.Version}");
                        Console.WriteLine($"활성 문서 수: {_wordApp.Documents.Count}");

                        if (_isTargetProg && !string.IsNullOrEmpty(_filePath))
                        {
                            // 기존 프로세스에서 원하는 파일 찾기
                            bool found = false;
                            foreach (Document doc in _wordApp.Documents)
                            {
                                if (string.Equals(doc.FullName, _filePath, StringComparison.OrdinalIgnoreCase))
                                {
                                    Console.WriteLine($"기존 프로세스에서 파일 찾음: {_filePath}");
                                    _document = doc;
                                    found = true;
                                    break;
                                }
                            }

                            // 파일을 찾지 못했다면 새로 열기
                            if (!found)
                            {
                                Console.WriteLine($"기존 프로세스에서 파일을 찾지 못해 새로 열기 시도: {_filePath}");
                                _document = _wordApp.Documents.Open(_filePath);
                                Console.WriteLine("파일 열기 성공");
                            }
                        }
                        else
                        {
                            // 활성 문서 가져오기
                            _document = _wordApp.ActiveDocument;
                            if (_document == null)
                            {
                                Console.WriteLine("활성 문서를 찾을 수 없습니다.");
                                return (string.Empty, new Dictionary<string, object>(), string.Empty);
                            }
                            Console.WriteLine($"활성 문서 이름: {_document.Name}");
                        }
                    }
                    else 
                    {
                        if (_isTargetProg)
                        {
                            Console.WriteLine("기존 프로세스가 없어, 새 Word 프로세스를 생성합니다.");
                            try
                            {
                                _wordApp = new WordApp();
                                Console.WriteLine("새 Word 애플리케이션 생성 성공");

                                Console.WriteLine($"활성 파일 경로: {_filePath}");

                                if (!string.IsNullOrEmpty(_filePath))
                                {
                                    Console.WriteLine($"파일 열기 시도: {_filePath}");
                                    _document = _wordApp.Documents.Open(_filePath);
                                    Console.WriteLine("파일 열기 성공");
                                }
                                Console.WriteLine($"활성 문서 상태: {(_document != null ? "존재함" : "없음")}");
                            }
                            catch (Exception createEx)
                            {
                                Console.WriteLine($"새 Word 애플리케이션 생성 실패: {createEx.Message}");
                                Console.WriteLine($"스택 트레이스: {createEx.StackTrace}");
                                throw;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"기존 Word 애플리케이션 찾기 실패: {ex.Message}");
                    Console.WriteLine("Word 애플리케이션이 실행 중이지 않습니다.");
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);
                }

                if (_wordApp == null)
                {
                    Console.WriteLine("Word 애플리케이션을 생성할 수 없습니다.");
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);
                }

                try
                {
                    Console.WriteLine("Word COM 객체 가져오기 시도...");
                    _wordApp = (WordApp)GetActiveObject("Word.Application");
                    Console.WriteLine("Word COM 객체 가져오기 성공");

                    Console.WriteLine("활성 문서 가져오기 시도...");
                    _document = _wordApp.ActiveDocument;

                    if (_document == null)
                    {
                        Console.WriteLine("활성 문서를 찾을 수 없습니다.");
                        return (string.Empty, new Dictionary<string, object>(), string.Empty);
                    }

                    Console.WriteLine($"활성 문서 정보:");
                    Console.WriteLine($"- 문서 이름: {_document.Name}");
                    Console.WriteLine($"- 전체 경로: {_document.FullName}");
                    Console.WriteLine($"- 저장 여부: {(_document.Saved ? "저장됨" : "저장되지 않음")}");
                    Console.WriteLine($"- 읽기 전용: {(_document.ReadOnly ? "예" : "아니오")}");

                    if(_isTargetProg)
                    {
                        Console.WriteLine("전체 문서 선택");
                        try 
                        {
                            _wordApp.Visible = true;
                            
                            // 전체 문서 선택
                            _document.Range().Select();
                            Console.WriteLine("전체 문서 선택 완료");
                            
                            // 선택된 내용을 클립보드에 복사
                            Console.WriteLine("클립보드에 복사 시도...");
                            _wordApp.Selection.Copy();
                            Console.WriteLine("클립보드 복사 완료");
                            
                            
                            // 클립보드 내용 확인
                            Console.WriteLine("클립보드 형식 확인 중...");
                            Console.WriteLine($"HTML 형식 존재: {Clipboard.ContainsText(TextDataFormat.Html)}");
                            Console.WriteLine($"일반 텍스트 존재: {Clipboard.ContainsText()}");
                            Console.WriteLine($"RTF 형식 존재: {Clipboard.ContainsText(TextDataFormat.Rtf)}");
                            
                            if (Clipboard.ContainsText(TextDataFormat.Html))
                            {
                                string htmlContent = Clipboard.GetText(TextDataFormat.Html);
                                Console.WriteLine($"HTML 데이터 길이: {htmlContent.Length}");
                                Console.WriteLine("HTML 데이터 일부: " + htmlContent.Substring(0, Math.Min(100, htmlContent.Length)));
                                return (htmlContent, new Dictionary<string, object>(), "전체 문서");
                            }
                            else
                            {
                                Console.WriteLine("클립보드에 HTML 형식 데이터 없음");
                                return (string.Empty, new Dictionary<string, object>(), string.Empty);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"클립보드 복사 중 오류 발생: {ex.Message}");
                            Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                            return (string.Empty, new Dictionary<string, object>(), string.Empty);
                        }
                    }
                    
                    var selection = _wordApp.Selection;
                    if (selection == null)
                    {
                        Console.WriteLine("선택된 텍스트가 없습니다.");
                        return (string.Empty, new Dictionary<string, object>(), string.Empty);
                    }

                    // 선택된 텍스트의 위치 정보 가져오기
                    string position = $"시작: {selection.Start}, 끝: {selection.End}";
                    Console.WriteLine($"선택된 텍스트 위치: {position}");

                    // 로그 윈도우의 컨텍스트 텍스트박스 업데이트
                    LogWindow.Instance.Dispatcher.Invoke(() =>
                    {
                        LogWindow.Instance.FilePathTextBox.Text = _document.FullName;
                        LogWindow.Instance.PositionTextBox.Text = position;
                        LogWindow.Instance.ContextTextBox.Text = selection.Text;
                    });

                    // HTML 형식으로 클립보드 복사 시도
                    try
                    {
                        selection.Copy();
                        if (Clipboard.ContainsText(TextDataFormat.Html))
                        {
                            string htmlContent = Clipboard.GetText(TextDataFormat.Html);

                            // 1) <!--StartFragment--> 와 <!--EndFragment--> 사이의 순수 HTML만 추출
                            int startIdx = htmlContent.IndexOf("<!--StartFragment-->");
                            int endIdx = htmlContent.IndexOf("<!--EndFragment-->");
                            if (startIdx != -1 && endIdx != -1 && endIdx > startIdx)
                            {
                                int fragContentStart = startIdx + "<!--StartFragment-->".Length;
                                int fragLength = endIdx - fragContentStart;
                                string rawFragment = htmlContent.Substring(fragContentStart, fragLength);

                                // 원본 HTML 출력
                                Console.WriteLine("=== 원본 HTML 시작 ===");
                                Console.WriteLine(rawFragment);
                                Console.WriteLine("=== 원본 HTML 끝 ===");

                                // 2) 불필요 태그/속성 제거 및 인접 노드 병합
                                //    + 최종적으로 줄바꿈/공백 정리
                                string cleanedHtml = CleanAndNormalizeHtml(rawFragment);

                                // test.html 파일 업데이트
                                try
                                {
                                    string htmlTemplate = @"<!DOCTYPE html>
<html lang=""en"">
<head>
    <meta charset=""UTF-8"">
    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">
    <title>텍스트 길이: {1}자</title>
</head>
<body>
{0}
</body>
</html>";

                                    string fullHtml = string.Format(htmlTemplate, cleanedHtml, cleanedHtml.Length);
                                    File.WriteAllText("test.html", fullHtml);
                                    Console.WriteLine("test.html 파일이 성공적으로 업데이트되었습니다.");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"test.html 파일 업데이트 실패: {ex.Message}");
                                }

                                // 로그 윈도우의 컨텍스트 텍스트박스 업데이트
                                LogWindow.Instance.Dispatcher.Invoke(() =>
                                {
                                    LogWindow.Instance.ContextTextBox.Text = cleanedHtml;
                                });

                                return (cleanedHtml, new Dictionary<string, object>(), position);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"HTML 클립보드 복사 실패: {ex.Message}");
                        return (string.Empty, new Dictionary<string, object>(), string.Empty);
                    }

                    return (string.Empty, new Dictionary<string, object>(), string.Empty);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Word COM 연결 오류: {ex.Message}");
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Word 데이터 읽기 오류 발생: {ex.Message}");
                Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                LogWindow.Instance.Log($"Word 데이터 읽기 오류: {ex.Message}");
                return (string.Empty, new Dictionary<string, object>(), string.Empty);
            }
            finally
            {
                if (_document != null) Marshal.ReleaseComObject(_document);
                if (_wordApp != null) Marshal.ReleaseComObject(_wordApp);
            }
        }

        /// <summary>
        /// 1) 불필요한 메타 태그/Office 네임스페이스 제거  
        /// 2) 빈 태그 스니펫 제거  
        /// 3) 인접 노드 병합  
        /// 4) lang, mso-* 속성 제거  
        /// 5) 빈 style 속성 제거  
        /// 6) 최종적으로 줄바꿈(\r,\n) 제거 및 태그 사이 공백 최소화
        /// </summary>
        private string CleanAndNormalizeHtml(string rawFragment)
        {
            // 1) HtmlAgilityPack으로 로드 (wrapper를 씌워 파싱)
            var htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.LoadHtml("<div id=\"wrapper\">" + rawFragment + "</div>");
            HtmlNode wrapper = htmlDoc.GetElementbyId("wrapper")!;

            // 2) 불필요한 메타/Office 전용 XML/네임스페이스/조건부 주석 제거
            RemoveUnwantedNodes(wrapper);

            // 3) style 내부에 표현된 Word 전용 속성과 lang 속성 제거
            RemoveWordSpecificAttributes(wrapper);

            // 4) class 속성 제거
            RemoveClassAttributes(wrapper);

            // 5) 빈 <span> 등 제거
            RemoveEmptySpans(wrapper);

            // 6) 인접한 <span> 병합
            MergeAdjacentSpans(wrapper);

            // 7) 빈 style 속성 제거 (위에서 이미 없는 style들은 지워졌을 것)
            RemoveEmptyStyleAttributes(wrapper);

            // 8) 이미지 데이터 분리 및 참조 처리
            ProcessImages(wrapper);

            // 9) 최종 HTML 문자열 추출
            string interimHtml = wrapper.InnerHtml;

            // 10) 보이지 않는 줄넘김(개행) 제거 및 태그 사이 공백 최소화
            string normalized = NormalizeWhitespace(interimHtml);

            return normalized;
        }

        /// <summary>
        /// 불필요한 메타, xml, Word 전용 네임스페이스, 조건부 주석 등을 삭제
        /// </summary>
        private void RemoveUnwantedNodes(HtmlNode root)
        {
            // 1) <meta> 태그 전부 제거
            var metas = root.SelectNodes("//meta");
            if (metas != null)
            {
                foreach (var meta in metas)
                    meta.Remove();
            }

            // 2) <xml> ... </xml> 노드 (Office 전용) 제거
            var xmlNodes = root.SelectNodes("//xml");
            if (xmlNodes != null)
            {
                foreach (var node in xmlNodes)
                    node.Remove();
            }

            // 3) Word 고유 네임스페이스가 들어간 모든 노드 제거 (예: <o:…>, <w:…>, <v:…>)
            var allNodes = root.SelectNodes("//*");
            if (allNodes != null)
            {
                foreach (var node in allNodes.ToList()) // ToList 로 복사 후 순회
                {
                    if (node.Name.StartsWith("o:", StringComparison.OrdinalIgnoreCase)
                        || node.Name.StartsWith("w:", StringComparison.OrdinalIgnoreCase)
                        || node.Name.StartsWith("v:", StringComparison.OrdinalIgnoreCase))
                    {
                        node.Remove();
                    }
                }
            }

            // 4) 조건부 주석(<!--[if gte mso ...]> ... <![endif]-->) 제거
            var comments = root.SelectNodes("//comment()");
            if (comments != null)
            {
                foreach (var commentNode in comments.Cast<HtmlCommentNode>())
                {
                    commentNode.Remove();
                }
            }
        }

        /// <summary>
        /// 모든 노드에서 Word 전용 'mso-...' 속성(스타일)과 lang 속성을 제거
        /// </summary>
        private void RemoveWordSpecificAttributes(HtmlNode root)
        {
            // 1) 모든 노드를 순회하면서 lang 속성을 제거
            var nodesWithLang = root.SelectNodes("//*[@lang]");
            if (nodesWithLang != null)
            {
                foreach (var node in nodesWithLang)
                {
                    node.Attributes.Remove("lang");
                }
            }

            // 2) style 속성이 있는 노드에서 mso-*, mso-fareast-language 등을 삭제
            var nodesWithStyle = root.SelectNodes("//*[@style]");
            if (nodesWithStyle != null)
            {
                foreach (var node in nodesWithStyle.ToList())
                {
                    var styleAttr = node.GetAttributeValue("style", "").Trim();
                    if (string.IsNullOrEmpty(styleAttr))
                    {
                        node.Attributes.Remove("style");
                        continue;
                    }

                    // 세미콜론(;)을 기준으로 개별 CSS 선언을 분리
                    var declarations = styleAttr
                        .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(decl => decl.Trim())
                        // mso-로 시작하는 모든 선언 제거
                        // .Where(decl => !decl.StartsWith("mso-", StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    if (declarations.Any())
                    {
                        // 나머지 유효 CSS 선언들을 다시 합친 뒤 속성 값으로 설정
                        string newStyleValue = string.Join(";", declarations) + ";";
                        node.SetAttributeValue("style", newStyleValue);
                    }
                    else
                    {
                        // 남은 style 선언이 없다면 아예 style 속성을 제거
                        node.Attributes.Remove("style");
                    }
                }
            }
        }

        /// <summary>
        /// 모든 노드에서 class 속성을 제거
        /// </summary>
        private void RemoveClassAttributes(HtmlNode root)
        {
            var nodesWithClass = root.SelectNodes("//*[@class]");
            if (nodesWithClass != null)
            {
                foreach (var node in nodesWithClass)
                {
                    node.Attributes.Remove("class");
                }
            }
        }

        /// <summary>
        /// style 속성이 비어 있거나 내용이 없는 <span> 태그 제거
        /// </summary>
        private void RemoveEmptySpans(HtmlNode root)
        {
            var spans = root.SelectNodes("//span");
            if (spans != null)
            {
                foreach (var span in spans.ToList())
                {
                    string styleAttr = span.GetAttributeValue("style", "").Trim();
                    string inner = span.InnerHtml.Trim();
                    if (string.IsNullOrEmpty(styleAttr) && string.IsNullOrEmpty(inner))
                    {
                        span.Remove();
                    }
                }
            }
        }

        /// <summary>
        /// 인접한 <span> 태그들 중 style 속성이 동일하면 병합
        /// </summary>
        private void MergeAdjacentSpans(HtmlNode root)
        {
            var parentNodes = root.SelectNodes("//*");
            if (parentNodes == null) return;

            foreach (var parent in parentNodes)
            {
                var children = parent.ChildNodes.ToList();
                for (int i = 0; i < children.Count - 1; i++)
                {
                    var curr = children[i];
                    var next = children[i + 1];

                    if (curr.Name.Equals("span", StringComparison.OrdinalIgnoreCase)
                        && next.Name.Equals("span", StringComparison.OrdinalIgnoreCase))
                    {
                        string styleCurr = curr.GetAttributeValue("style", "");
                        string styleNext = next.GetAttributeValue("style", "");
                        if (styleCurr == styleNext)
                        {
                            curr.InnerHtml = curr.InnerHtml + next.InnerHtml;
                            next.Remove();
                            // 노드 리스트 갱신 후 인덱스 재조정
                            children = parent.ChildNodes.ToList();
                            i--;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 인라인 style="" 속성이 비어 있으면 해당 속성 제거
        /// </summary>
        private void RemoveEmptyStyleAttributes(HtmlNode root)
        {
            var nodesWithStyle = root.SelectNodes("//*[@style]");
            if (nodesWithStyle != null)
            {
                foreach (var node in nodesWithStyle.ToList())
                {
                    var val = node.GetAttributeValue("style", "").Trim();
                    if (string.IsNullOrEmpty(val))
                        node.Attributes.Remove("style");
                }
            }
        }

        /// <summary>
        /// 이미지 데이터를 분리하고 참조로 대체
        /// </summary>
        private void ProcessImages(HtmlNode root)
        {
            // 이미지 저장 디렉토리 생성
            string imageDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "images");
            if (!Directory.Exists(imageDir))
            {
                Directory.CreateDirectory(imageDir);
            }

            // 모든 이미지 노드 찾기
            var images = root.SelectNodes("//img");
            if (images != null)
            {
                foreach (var img in images)
                {
                    string src = img.GetAttributeValue("src", "");
                    if (src.StartsWith("data:image"))
                    {
                        try
                        {
                            // Base64 데이터 추출
                            string[] parts = src.Split(',');
                            if (parts.Length > 1)
                            {
                                string imageData = parts[1];
                                string imageId = Guid.NewGuid().ToString();
                                string imagePath = Path.Combine(imageDir, $"{imageId}.jpg");

                                // 이미지 데이터를 파일로 저장
                                byte[] imageBytes = Convert.FromBase64String(imageData);
                                File.WriteAllBytes(imagePath, imageBytes);

                                // 절대 경로로 이미지 참조
                                string absolutePath = Path.GetFullPath(imagePath);
                                img.SetAttributeValue("src", absolutePath);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"이미지 처리 중 오류 발생: {ex.Message}");
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 1) 보이지 않는 줄바꿈(\r, \n) 전부 제거  
        /// 2) 태그 사이 공백(스페이스/탭/개행) 제거: ">   <" → "><"  
        /// 3) 연속된 공백(스페이스) 2개 이상 → 1개로 축소  
        /// </summary>
        private string NormalizeWhitespace(string html)
        {
            if (string.IsNullOrEmpty(html))
                return string.Empty;

            // 원본 공백 유지
            return html.Trim();
        }

        public override (ulong? FileId, uint? VolumeId, string FileType, string FileName, string FilePath) GetFileInfo()
        {
            WordApp? tempWordApp = null;
            Document? tempDocument = null;
            
            try
            {
                Console.WriteLine("Word COM 객체 가져오기 시도...");
                tempWordApp = (WordApp)GetActiveObject("Word.Application");
                Console.WriteLine("Word COM 객체 가져오기 성공");

                if (_isTargetProg && !string.IsNullOrEmpty(_filePath))
                {
                    // 모든 문서 확인
                    foreach (Document doc in tempWordApp.Documents)
                    {
                        try
                        {
                            string filePath = doc.FullName;
                            string fileName = doc.Name;
                            
                            Console.WriteLine($"Word 문서 정보:");
                            Console.WriteLine($"- 파일 경로: {filePath}");
                            Console.WriteLine($"- 파일 이름: {fileName}");
                            
                            if (string.IsNullOrEmpty(filePath))
                            {
                                Console.WriteLine("파일 경로가 비어있습니다.");
                                continue;
                            }
                            
                            if (string.Equals(_filePath, filePath, StringComparison.OrdinalIgnoreCase))
                            {
                                var fileIdInfo = GetFileId(filePath);
                                
                                if (fileIdInfo == null)
                                {
                                    Console.WriteLine("파일 ID 정보를 가져오지 못했습니다.");
                                }
                                else
                                {
                                    Console.WriteLine($"파일 ID 정보:");
                                    Console.WriteLine($"- FileId: {fileIdInfo.Value.FileId}");
                                    Console.WriteLine($"- VolumeId: {fileIdInfo.Value.VolumeId}");
                                }
                                
                                return (
                                    fileIdInfo?.FileId,
                                    fileIdInfo?.VolumeId,
                                    "Word",
                                    fileName,
                                    filePath
                                );
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"문서 처리 중 오류: {ex.Message}");
                            continue;
                        }
                    }
                }
                else
                {
                    // 활성 문서 정보 가져오기
                    tempDocument = tempWordApp.ActiveDocument;
                    if (tempDocument != null)
                    {
                        string filePath = tempDocument.FullName;
                        string fileName = tempDocument.Name;
                        
                        Console.WriteLine($"활성 Word 문서 정보:");
                        Console.WriteLine($"- 파일 경로: {filePath}");
                        Console.WriteLine($"- 파일 이름: {fileName}");
                        
                        var fileIdInfo = GetFileId(filePath);
                        
                        if (fileIdInfo == null)
                        {
                            Console.WriteLine("파일 ID 정보를 가져오지 못했습니다.");
                        }
                        else
                        {
                            Console.WriteLine($"파일 ID 정보:");
                            Console.WriteLine($"- FileId: {fileIdInfo.Value.FileId}");
                            Console.WriteLine($"- VolumeId: {fileIdInfo.Value.VolumeId}");
                        }
                        
                        return (
                            fileIdInfo?.FileId,
                            fileIdInfo?.VolumeId,
                            "Word",
                            fileName,
                            filePath
                        );
                    }
                }
                
                Console.WriteLine("문서를 찾을 수 없습니다.");
                return (null, null, "Word", string.Empty, string.Empty);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"파일 정보 가져오기 오류: {ex.Message}");
                Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                return (null, null, "Word", string.Empty, string.Empty);
            }
            finally
            {
                if (tempDocument != null) Marshal.ReleaseComObject(tempDocument);
                if (tempWordApp != null) Marshal.ReleaseComObject(tempWordApp);
            }
        }
    }
}
