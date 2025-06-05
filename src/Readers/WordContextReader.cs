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

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

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

        public override (string SelectedText, Dictionary<string, object> StyleAttributes, string LineNumber) GetSelectedTextWithStyle()
{
    try
    {
        Console.WriteLine("Word 데이터 읽기 시작...");

        var wordProcesses = Process.GetProcessesByName("WINWORD");
        if (wordProcesses.Length == 0)
        {
            Console.WriteLine("실행 중인 Word 애플리케이션을 찾을 수 없습니다.");
            throw new InvalidOperationException("Word is not running");
        }

        Process? activeWordProcess = null;
        foreach (var process in wordProcesses)
        {
            if (process.MainWindowHandle != IntPtr.Zero && process.MainWindowTitle.Length > 0)
            {
                Console.WriteLine($"Word 프로세스 정보:");
                Console.WriteLine($"- 프로세스 ID: {process.Id}");
                Console.WriteLine($"- 창 제목: {process.MainWindowTitle}");
                Console.WriteLine($"- 실행 경로: {process.MainModule?.FileName}");

                if (process.MainWindowHandle == GetForegroundWindow())
                {
                    activeWordProcess = process;
                    Console.WriteLine("이 Word 창이 현재 활성화되어 있습니다.");
                }
            }
        }

        if (activeWordProcess == null)
        {
            Console.WriteLine("활성화된 Word 창을 찾을 수 없습니다.");
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

            var selection = _wordApp.Selection;
            if (selection == null)
            {
                Console.WriteLine("선택된 텍스트가 없습니다.");
                return (string.Empty, new Dictionary<string, object>(), string.Empty);
            }

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

                        // 2) 불필요 태그/속성 제거 및 인접 노드 병합
                        //    + 최종적으로 줄바꿈/공백 정리
                        string cleanedHtml = CleanAndNormalizeHtml(rawFragment);

                        return (cleanedHtml, new Dictionary<string, object>(), string.Empty);
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

    // 4) 빈 <span> 등 제거
    RemoveEmptySpans(wrapper);

    // 5) 인접한 <span> 병합
    MergeAdjacentSpans(wrapper);

    // 6) 빈 style 속성 제거 (위에서 이미 없는 style들은 지워졌을 것)
    RemoveEmptyStyleAttributes(wrapper);

    // 7) 최종 HTML 문자열 추출
    string interimHtml = wrapper.InnerHtml;

    // 8) 보이지 않는 줄넘김(개행) 제거 및 태그 사이 공백 최소화
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
            string commentText = commentNode.Comment;
            if (commentText.TrimStart().StartsWith("[if ", StringComparison.OrdinalIgnoreCase))
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
                // mso-bidi-font-family, mso-fareast-language, mso-*로 시작되는 선언 제거
                .Where(decl =>
                    !decl.StartsWith("mso-bidi-font-family", StringComparison.OrdinalIgnoreCase)
                    && !decl.StartsWith("mso-fareast-language", StringComparison.OrdinalIgnoreCase)
                    && !decl.StartsWith("mso-rtl-language", StringComparison.OrdinalIgnoreCase)
                    && !decl.StartsWith("mso-ascii-font-family", StringComparison.OrdinalIgnoreCase)
                    && !decl.StartsWith("mso-ascii-theme-font", StringComparison.OrdinalIgnoreCase)
                    && !decl.StartsWith("mso-fareast-theme-font", StringComparison.OrdinalIgnoreCase)
                    && !decl.StartsWith("mso-hansi-font-family", StringComparison.OrdinalIgnoreCase)
                    && !decl.StartsWith("mso-hansi-theme-font", StringComparison.OrdinalIgnoreCase)
                    // 필요 시 추가로 mso- 로 시작하는 다른 항목들을 이곳에 더할 수 있습니다.
                    )
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
/// 1) 보이지 않는 줄바꿈(\r, \n) 전부 제거  
/// 2) 태그 사이 공백(스페이스/탭/개행) 제거: ">   <" → "><"  
/// 3) 연속된 공백(스페이스) 2개 이상 → 1개로 축소  
/// </summary>
private string NormalizeWhitespace(string html)
{
    if (string.IsNullOrEmpty(html))
        return string.Empty;

    // 1) 연속된 공백 2개 이상 → 1개로 축소 (단, 태그 내부의 공백은 보존)
    string collapsedSpaces = Regex.Replace(html, @"(?<=[^>]) {2,}(?=[^<])", " ");

    // 2) 태그 사이의 불필요한 공백만 제거 (줄바꿈은 보존)
    string cleanedHtml = Regex.Replace(collapsedSpaces, @">\s+<", "><");

    // 3) 텍스트 내용의 공백 보존
    cleanedHtml = Regex.Replace(cleanedHtml, @"(?<=>)(\s+)(?=<)", match => match.Value);

    return cleanedHtml.Trim();
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

                Console.WriteLine("활성 문서 가져오기 시도...");
                tempDocument = tempWordApp.ActiveDocument;
                
                if (tempDocument == null)
                {
                    Console.WriteLine("활성 문서를 찾을 수 없습니다.");
                    return (null, null, "Word", string.Empty, string.Empty);
                }

                string filePath = tempDocument.FullName;
                string fileName = tempDocument.Name;
                
                Console.WriteLine($"Word 문서 정보:");
                Console.WriteLine($"- 파일 경로: {filePath}");
                Console.WriteLine($"- 파일 이름: {fileName}");
                
                if (string.IsNullOrEmpty(filePath))
                {
                    Console.WriteLine("파일 경로가 비어있습니다.");
                    return (null, null, "Word", fileName, string.Empty);
                }
                
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
