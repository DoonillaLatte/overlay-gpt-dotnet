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

namespace overlay_gpt
{
    public class WordContextReader : BaseContextReader
    {
        private Application? _wordApp;
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

        private string GetStyledText(string text, Dictionary<string, object> styleAttributes)
        {
            string result = text;
            if (styleAttributes.ContainsKey("UnderlineStyle") && styleAttributes["UnderlineStyle"]?.ToString() == "Single")
                result = $"<u>{result}</u>";
            if (styleAttributes.ContainsKey("FontWeight") && styleAttributes["FontWeight"]?.ToString() == "Bold")
                result = $"<b>{result}</b>";
            if (styleAttributes.ContainsKey("FontItalic") && Convert.ToBoolean(styleAttributes["FontItalic"]))
                result = $"<i>{result}</i>";
            if (styleAttributes.ContainsKey("FontStrikethrough") && Convert.ToBoolean(styleAttributes["FontStrikethrough"]))
                result = $"<s>{result}</s>";
            return result;
        }

        private int ConvertColorToRGB(int bgrColor)
        {
            int r = bgrColor & 0xFF;
            int g = (bgrColor >> 8) & 0xFF;
            int b = (bgrColor >> 16) & 0xFF;
            return (r << 16) | (g << 8) | b;
        }

        private int GetHighlightColorRGB(int highlightColor)
        {
            // 사용자 정의 RGB 색상인 경우 (0xFFFFFF보다 큰 값)
            if (highlightColor > 0xFFFFFF)
            {
                return ConvertColorToRGB(highlightColor);
            }

            // WdColorIndex 열거형 값에 따른 RGB 색상 매핑
            switch (highlightColor)
            {
                case (int)WdColorIndex.wdYellow: return 0xFFFF00;  // 노랑
                case (int)WdColorIndex.wdBrightGreen: return 0x00FF00;  // 밝은 초록
                case (int)WdColorIndex.wdTurquoise: return 0x00FFFF;  // 청록
                case (int)WdColorIndex.wdPink: return 0xFF00FF;  // 분홍
                case (int)WdColorIndex.wdBlue: return 0x0000FF;  // 파랑
                case (int)WdColorIndex.wdRed: return 0xFF0000;  // 빨강
                case (int)WdColorIndex.wdDarkBlue: return 0x000080;  // 진한 파랑
                case (int)WdColorIndex.wdTeal: return 0x008080;  // 청녹
                case (int)WdColorIndex.wdGreen: return 0x008000;  // 초록
                case (int)WdColorIndex.wdViolet: return 0x800080;  // 보라
                case (int)WdColorIndex.wdDarkRed: return 0x800000;  // 진한 빨강
                case (int)WdColorIndex.wdDarkYellow: return 0x808000;  // 진한 노랑
                case (int)WdColorIndex.wdGray50: return 0x808080;  // 회색
                case (int)WdColorIndex.wdGray25: return 0xC0C0C0;  // 연한 회색
                default: return 0xFFFFFF;  // 기본값 (흰색)
            }
        }

        private string GetTextStyleString(Dictionary<string, object> styleAttributes)
        {
            var styleList = new List<string>();
            
            if (styleAttributes.ContainsKey("FontName"))
            {
                string fontName = styleAttributes["FontName"]?.ToString() ?? string.Empty;
                if (!string.IsNullOrEmpty(fontName) && fontName != "Calibri" && fontName != "Arial" && fontName != "맑은 고딕")
                {
                    styleList.Add($"font-family: {fontName}");
                }
            }
            
            if (styleAttributes.ContainsKey("FontSize"))
            {
                double fontSize = Convert.ToDouble(styleAttributes["FontSize"]);
                if (fontSize != 11)
                {
                    styleList.Add($"font-size: {fontSize}pt");
                }
            }

            if (styleAttributes.ContainsKey("ForegroundColor"))
            {
                int fgColor = Convert.ToInt32(styleAttributes["ForegroundColor"]);
                if (fgColor != 0)
                {
                    int rgbColor = ConvertColorToRGB(fgColor);
                    string hexColor = $"#{rgbColor:X6}";
                    styleList.Add($"color: {hexColor}");
                }
            }

            if (styleAttributes.ContainsKey("HighlightColor"))
            {
                int highlightColor = Convert.ToInt32(styleAttributes["HighlightColor"]);
                if (highlightColor != 0)
                {
                    int rgbColor = GetHighlightColorRGB(highlightColor);
                    string hexColor = $"#{rgbColor:X6}";
                    styleList.Add($"background-color: {hexColor}");
                }
            }

            return string.Join("; ", styleList);
        }

        public override (string SelectedText, Dictionary<string, object> StyleAttributes, string LineNumber) GetSelectedTextWithStyle(bool readAllContent = false)
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
                    _wordApp = (Application)GetActiveObject("Word.Application");
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

                    var range = readAllContent ? _document.Content : _wordApp.Selection.Range;
                    if (range == null)
                    {
                        Console.WriteLine("선택된 텍스트가 없습니다.");
                        return (string.Empty, new Dictionary<string, object>(), string.Empty);
                    }

                    string selectedText = range.Text;
                    var styleAttributes = new Dictionary<string, object>();
                    var styledTextBuilder = new StringBuilder();

                    // 전체 문서의 라인 수를 기준으로 선택된 텍스트의 시작과 끝 라인 번호 계산
                    int totalLines = _document.ComputeStatistics(WdStatistic.wdStatisticLines);
                    int startLine = 1;
                    int endLine = totalLines;

                    // 선택된 텍스트의 시작과 끝 위치를 기준으로 라인 번호 계산
                    int selectionStart = range.Start;
                    int selectionEnd = range.End;

                    for (int i = 1; i <= totalLines; i++)
                    {
                        var lineRange = _document.Range(_document.GoTo(WdGoToItem.wdGoToLine, WdGoToDirection.wdGoToAbsolute, i).Start,
                                                      _document.GoTo(WdGoToItem.wdGoToLine, WdGoToDirection.wdGoToAbsolute, i).End);
                        
                        if (lineRange.Start <= selectionStart && lineRange.End >= selectionStart)
                        {
                            startLine = i;
                        }
                        if (lineRange.Start <= selectionEnd && lineRange.End >= selectionEnd)
                        {
                            endLine = i;
                            break;
                        }
                    }

                    string lineNumber = $"{startLine}-{endLine}";

                    // 선택된 텍스트의 각 부분에 대해 스타일 정보 수집
                    var start = range.Start;
                    var end = range.End;

                    Dictionary<string, object>? currentStyle = null;
                    var currentTextBuilder = new StringBuilder();

                    for (int i = start; i < end; i++)
                    {
                        var charRange = _document.Range(i, i + 1);
                        var comFont = (WordFont)charRange.Font;
                        var charStyle = new Dictionary<string, object>
                        {
                            ["FontName"] = comFont.Name,
                            ["FontSize"] = comFont.Size,
                            ["FontWeight"] = comFont.Bold == -1 ? "Bold" : "Normal",
                            ["FontItalic"] = comFont.Italic == -1,
                            ["UnderlineStyle"] = comFont.Underline == WdUnderline.wdUnderlineSingle ? "Single" : "None",
                            ["ForegroundColor"] = comFont.Color,
                            ["HighlightColor"] = charRange.HighlightColorIndex
                        };

                        string currentChar = charRange.Text;

                        // 줄넘김 처리
                        if (currentChar == "\r" || currentChar == "\n")
                        {
                            // 현재까지의 텍스트를 스타일과 함께 추가
                            if (currentStyle != null && currentTextBuilder.Length > 0)
                            {
                                string styledText = GetStyledText(currentTextBuilder.ToString(), currentStyle);
                                string styleString = GetTextStyleString(currentStyle);
                                if (!string.IsNullOrEmpty(styleString))
                                {
                                    styledText = $"<span style='{styleString}'>{styledText}</span>";
                                }
                                styledTextBuilder.Append(styledText);
                                currentTextBuilder.Clear();
                            }
                            styledTextBuilder.Append("<br>");
                            continue;
                        }

                        if (currentStyle == null)
                        {
                            currentStyle = charStyle;
                            currentTextBuilder.Append(currentChar);
                        }
                        else if (AreStylesEqual(currentStyle, charStyle))
                        {
                            currentTextBuilder.Append(currentChar);
                        }
                        else
                        {
                            // 현재까지의 텍스트를 스타일과 함께 추가
                            string styledText = GetStyledText(currentTextBuilder.ToString(), currentStyle);
                            string styleString = GetTextStyleString(currentStyle);
                            if (!string.IsNullOrEmpty(styleString))
                            {
                                styledText = $"<span style='{styleString}'>{styledText}</span>";
                            }
                            styledTextBuilder.Append(styledText);

                            // 새로운 스타일로 시작
                            currentStyle = charStyle;
                            currentTextBuilder.Clear();
                            currentTextBuilder.Append(currentChar);
                        }
                    }

                    // 마지막 텍스트 처리
                    if (currentStyle != null && currentTextBuilder.Length > 0)
                    {
                        string styledText = GetStyledText(currentTextBuilder.ToString(), currentStyle);
                        string styleString = GetTextStyleString(currentStyle);
                        if (!string.IsNullOrEmpty(styleString))
                        {
                            styledText = $"<span style='{styleString}'>{styledText}</span>";
                        }
                        styledTextBuilder.Append(styledText);
                    }

                    return (styledTextBuilder.ToString(), styleAttributes, lineNumber);
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

        public override (ulong? FileId, uint? VolumeId, string FileType, string FileName, string FilePath) GetFileInfo()
        {
            Application? tempWordApp = null;
            Document? tempDocument = null;
            
            try
            {
                Console.WriteLine("Word COM 객체 가져오기 시도...");
                tempWordApp = (Application)GetActiveObject("Word.Application");
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

        private bool AreStylesEqual(Dictionary<string, object> style1, Dictionary<string, object> style2)
        {
            if (style1.Count != style2.Count) return false;

            foreach (var key in style1.Keys)
            {
                if (!style2.ContainsKey(key)) return false;
                if (!style1[key].Equals(style2[key])) return false;
            }

            return true;
        }
    }
}
