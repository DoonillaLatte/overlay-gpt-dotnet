using System;
using System.Collections.Generic;
using System.Windows.Automation;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;  // Office Core 15.0 네임스페이스
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.IO;

namespace overlay_gpt
{
    public class PPTContextReader : BaseContextReader
    {
        private Application? _pptApp;
        private Presentation? _presentation;
        private Slide? _slide;

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
            if (styleAttributes.ContainsKey("FontItalic") && styleAttributes["FontItalic"] is MsoTriState italic && italic == MsoTriState.msoTrue)
                result = $"<i>{result}</i>";
            if (styleAttributes.ContainsKey("FontStrikethrough") && styleAttributes["FontStrikethrough"] is MsoTriState strikethrough && strikethrough == MsoTriState.msoTrue)
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

            return string.Join("; ", styleList);
        }

        public override (string SelectedText, Dictionary<string, object> StyleAttributes, string LineNumber) GetSelectedTextWithStyle()
        {
            try
            {
                Console.WriteLine("PowerPoint 데이터 읽기 시작...");

                var pptProcesses = Process.GetProcessesByName("POWERPNT");
                if (pptProcesses.Length == 0)
                {
                    Console.WriteLine("실행 중인 PowerPoint 애플리케이션을 찾을 수 없습니다.");
                    throw new InvalidOperationException("PowerPoint is not running");
                }

                Process? activePPTProcess = null;
                foreach (var process in pptProcesses)
                {
                    if (process.MainWindowHandle != IntPtr.Zero && process.MainWindowTitle.Length > 0)
                    {
                        Console.WriteLine($"PowerPoint 프로세스 정보:");
                        Console.WriteLine($"- 프로세스 ID: {process.Id}");
                        Console.WriteLine($"- 창 제목: {process.MainWindowTitle}");
                        Console.WriteLine($"- 실행 경로: {process.MainModule?.FileName}");
                        
                        if (process.MainWindowHandle == GetForegroundWindow())
                        {
                            activePPTProcess = process;
                            Console.WriteLine("이 PowerPoint 창이 현재 활성화되어 있습니다.");
                        }
                    }
                }

                if (activePPTProcess == null)
                {
                    Console.WriteLine("활성화된 PowerPoint 창을 찾을 수 없습니다.");
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);
                }

                try
                {
                    Console.WriteLine("PowerPoint COM 객체 가져오기 시도...");
                    _pptApp = (Application)GetActiveObject("PowerPoint.Application");
                    Console.WriteLine("PowerPoint COM 객체 가져오기 성공");

                    Console.WriteLine("활성 프레젠테이션 가져오기 시도...");
                    _presentation = _pptApp.ActivePresentation;
                    
                    if (_presentation == null)
                    {
                        Console.WriteLine("활성 프레젠테이션을 찾을 수 없습니다.");
                        return (string.Empty, new Dictionary<string, object>(), string.Empty);
                    }
                    
                    Console.WriteLine($"활성 프레젠테이션 정보:");
                    Console.WriteLine($"- 프레젠테이션 이름: {_presentation.Name}");
                    Console.WriteLine($"- 전체 경로: {_presentation.FullName}");
                    Console.WriteLine($"- 저장 여부: {(_presentation.Saved == MsoTriState.msoTrue ? "저장됨" : "저장되지 않음")}");
                    Console.WriteLine($"- 읽기 전용: {(_presentation.ReadOnly == MsoTriState.msoTrue ? "예" : "아니오")}");

                    _slide = _pptApp.ActiveWindow?.View?.Slide;
                    if (_slide == null)
                    {
                        Console.WriteLine("활성 슬라이드를 찾을 수 없습니다.");
                        return (string.Empty, new Dictionary<string, object>(), string.Empty);
                    }

                    var selection = _pptApp.ActiveWindow?.Selection;
                    if (selection == null)
                    {
                        Console.WriteLine("선택된 텍스트가 없습니다.");
                        return (string.Empty, new Dictionary<string, object>(), string.Empty);
                    }

                    var styledTextBuilder = new StringBuilder();
                    var styleAttributes = new Dictionary<string, object>();

                    foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in selection.ShapeRange)
                    {
                        if (shape.HasTextFrame == MsoTriState.msoTrue)
                        {
                            var textFrame = shape.TextFrame;
                            var textRange = textFrame.TextRange;
                            
                            // 텍스트 스타일 정보 수집
                            var font = textRange.Font;
                            styleAttributes["FontName"] = font.Name;
                            styleAttributes["FontSize"] = font.Size;
                            styleAttributes["FontWeight"] = font.Bold == MsoTriState.msoTrue ? "Bold" : "Normal";
                            styleAttributes["FontItalic"] = font.Italic;
                            styleAttributes["ForegroundColor"] = font.Color.RGB;

                            string text = textRange.Text;
                            string styledText = GetStyledText(text, styleAttributes);
                            string styleString = GetTextStyleString(styleAttributes);
                            
                            if (!string.IsNullOrEmpty(styleString))
                            {
                                styledText = $"<span style='{styleString}'>{styledText}</span>";
                            }
                            
                            styledTextBuilder.Append(styledText);
                        }
                    }

                    string selectedText = styledTextBuilder.ToString();
                    string lineNumber = $"Slide {_slide.SlideIndex}";

                    return (selectedText, styleAttributes, lineNumber);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"PowerPoint COM 연결 오류: {ex.Message}");
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"PowerPoint 데이터 읽기 오류 발생: {ex.Message}");
                Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                LogWindow.Instance.Log($"PowerPoint 데이터 읽기 오류: {ex.Message}");
                return (string.Empty, new Dictionary<string, object>(), string.Empty);
            }
            finally
            {
                if (_slide != null) Marshal.ReleaseComObject(_slide);
                if (_presentation != null) Marshal.ReleaseComObject(_presentation);
                if (_pptApp != null) Marshal.ReleaseComObject(_pptApp);
            }
        }

        public override (ulong? FileId, uint? VolumeId, string FileType, string FileName, string FilePath) GetFileInfo()
        {
            try
            {
                if (_presentation == null)
                    return (null, null, "PowerPoint", string.Empty, string.Empty);

                string filePath = _presentation.FullName;
                if (string.IsNullOrEmpty(filePath))
                    return (null, null, "PowerPoint", _presentation.Name, string.Empty);

                var fileIdInfo = GetFileId(filePath);
                return (
                    fileIdInfo?.FileId,
                    fileIdInfo?.VolumeId,
                    "PowerPoint",
                    _presentation.Name,
                    filePath
                );
            }
            catch (Exception ex)
            {
                Console.WriteLine($"파일 정보 가져오기 오류: {ex.Message}");
                return (null, null, "PowerPoint", string.Empty, string.Empty);
            }
        }
    }
}
