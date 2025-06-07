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
using Forms = System.Windows.Forms;

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
            if (styleAttributes.ContainsKey("FontSize"))
            {
                double fontSize = Convert.ToDouble(styleAttributes["FontSize"]);
                result = $"<span style='font-size: {fontSize}pt'>{result}</span>";
            }
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

            // 전경색 (텍스트 색상)
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

            // 배경색
            if (styleAttributes.ContainsKey("BackgroundColor"))
            {
                int bgColor = Convert.ToInt32(styleAttributes["BackgroundColor"]);
                if (bgColor != 0)
                {
                    int rgbColor = ConvertColorToRGB(bgColor);
                    double transparency = styleAttributes.ContainsKey("BackgroundTransparency") ? 
                        Convert.ToDouble(styleAttributes["BackgroundTransparency"]) : 0.0;
                    
                    if (transparency < 1.0)
                    {
                        int r = (rgbColor >> 16) & 0xFF;
                        int g = (rgbColor >> 8) & 0xFF;
                        int b = rgbColor & 0xFF;
                        styleList.Add($"background-color: rgba({r}, {g}, {b}, {1.0 - transparency})");
                    }
                }
            }

            // 하이라이트 색
            if (styleAttributes.ContainsKey("HighlightColor"))
            {
                int highlightColor = Convert.ToInt32(styleAttributes["HighlightColor"]);
                if (highlightColor != 0)
                {
                    int rgbColor = ConvertColorToRGB(highlightColor);
                    int r = (rgbColor >> 16) & 0xFF;
                    int g = (rgbColor >> 8) & 0xFF;
                    int b = rgbColor & 0xFF;
                    styleList.Add($"background-color: rgb({r}, {g}, {b})");
                }
            }

            // 텍스트 정렬 설정
            string textAlign = "left";
            if (styleAttributes.ContainsKey("TextAlign"))
            {
                textAlign = styleAttributes["TextAlign"]?.ToString() ?? "left";
            }
            styleList.Add($"text-align: {textAlign}");

            // 수직 정렬 설정
            string verticalAlign = "top";
            if (styleAttributes.ContainsKey("VerticalAlign"))
            {
                verticalAlign = styleAttributes["VerticalAlign"]?.ToString() ?? "top";
            }
            styleList.Add($"vertical-align: {verticalAlign}");

            // Flexbox 속성 설정
            styleList.Add("display: flex");
            styleList.Add("align-items: center");
            
            // 수평 정렬에 따른 justify-content 설정
            switch (textAlign)
            {
                case "center":
                    styleList.Add("justify-content: center");
                    break;
                case "right":
                    styleList.Add("justify-content: flex-end");
                    break;
                default:
                    styleList.Add("justify-content: flex-start");
                    break;
            }

            // 수직 정렬에 따른 align-items 설정
            switch (verticalAlign)
            {
                case "middle":
                    styleList.Add("align-items: center");
                    break;
                case "bottom":
                    styleList.Add("align-items: flex-end");
                    break;
                default:
                    styleList.Add("align-items: flex-start");
                    break;
            }

            return string.Join("; ", styleList);
        }

        private string GetShapeStyleString(Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            var styleList = new List<string>();
            
            // 위치와 크기
            styleList.Add($"position: absolute");
            styleList.Add($"left: {shape.Left}px");
            styleList.Add($"top: {shape.Top}px");
            styleList.Add($"width: {shape.Width}px");
            styleList.Add($"height: {shape.Height}px");
            
            // 회전
            if (shape.Rotation != 0)
            {
                styleList.Add($"transform: rotate({shape.Rotation}deg)");
            }

            // 투명도
            if (shape.Fill.Visible == MsoTriState.msoTrue)
            {
                if (shape.Fill.Type == MsoFillType.msoFillGradient)
                {
                    // 그라데이션 처리
                    var gradientStops = new List<string>();
                    for (int i = 1; i <= shape.Fill.GradientStops.Count; i++)
                    {
                        var stop = shape.Fill.GradientStops[i];
                        int rgbColor = ConvertColorToRGB(stop.Color.RGB);
                        string hexColor = $"#{rgbColor:X6}";
                        gradientStops.Add($"{hexColor} {stop.Position * 100}%");
                    }
                    string gradientDirection = shape.Fill.GradientAngle == 0 ? "to right" : 
                                            shape.Fill.GradientAngle == 90 ? "to bottom" :
                                            shape.Fill.GradientAngle == 180 ? "to left" :
                                            "to top";
                    styleList.Add($"background: linear-gradient({gradientDirection}, {string.Join(", ", gradientStops)})");
                }
                else if (shape.Fill.ForeColor.RGB != 0)
                {
                    int rgbColor = ConvertColorToRGB(shape.Fill.ForeColor.RGB);
                    double transparency = shape.Fill.Transparency;
                    if (transparency < 1.0)
                    {
                        int r = (rgbColor >> 16) & 0xFF;
                        int g = (rgbColor >> 8) & 0xFF;
                        int b = rgbColor & 0xFF;
                        styleList.Add($"background-color: rgba({r}, {g}, {b}, {1.0 - transparency})");
                    }
                }
            }
            
            // 테두리
            if (shape.Line.Visible == MsoTriState.msoTrue)
            {
                int borderColor = ConvertColorToRGB(shape.Line.ForeColor.RGB);
                string borderHexColor = $"#{borderColor:X6}";
                string borderStyle = "solid"; // 기본값
                if (shape.Line.DashStyle != MsoLineDashStyle.msoLineSolid)
                {
                    borderStyle = "dashed";
                }
                styleList.Add($"border: {shape.Line.Weight}px {borderStyle} {borderHexColor}");
            }

            // 그림자 효과
            if (shape.Shadow.Visible == MsoTriState.msoTrue)
            {
                int shadowColor = ConvertColorToRGB(shape.Shadow.ForeColor.RGB);
                string shadowHexColor = $"#{shadowColor:X6}";
                double shadowBlur = shape.Shadow.Size;
                double shadowOffsetX = shape.Shadow.OffsetX;
                double shadowOffsetY = shape.Shadow.OffsetY;
                double shadowTransparency = shape.Shadow.Transparency / 100.0;
                styleList.Add($"box-shadow: {shadowOffsetX}px {shadowOffsetY}px {shadowBlur}px rgba({shadowColor >> 16 & 0xFF}, {shadowColor >> 8 & 0xFF}, {shadowColor & 0xFF}, {1 - shadowTransparency})");
            }

            // 모서리 둥글기
            if (shape.AutoShapeType != MsoAutoShapeType.msoShapeRectangle)
            {
                // 자동 도형의 경우 모서리 둥글기 적용
                styleList.Add($"border-radius: {Math.Min(shape.Width, shape.Height) * 0.1}px");
            }

            // 3D 효과
            if (shape.ThreeD.Visible == MsoTriState.msoTrue)
            {
                styleList.Add($"transform-style: preserve-3d");
                styleList.Add($"perspective: 1000px");
                styleList.Add($"transform: rotateX({shape.ThreeD.RotationX}deg) rotateY({shape.ThreeD.RotationY}deg)");
            }

            // Z-인덱스 (레이어 순서)
            styleList.Add($"z-index: {shape.ZOrderPosition}");

            return string.Join("; ", styleList);
        }

        private string GetShapeType(Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            switch (shape.Type)
            {
                case MsoShapeType.msoAutoShape:
                    return "div";
                case MsoShapeType.msoPicture:
                    return "img";
                case MsoShapeType.msoTextBox:
                    return "div";
                case MsoShapeType.msoLine:
                    return "div";
                case MsoShapeType.msoChart:
                    return "div";
                case MsoShapeType.msoTable:
                    return "table";
                case MsoShapeType.msoSmartArt:
                    return "div";
                default:
                    return "div";
            }
        }

        private string ConvertShapeToHtml(Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            string shapeType = GetShapeType(shape);
            string styleString = GetShapeStyleString(shape);
            string content = string.Empty;

            if (shape.HasTextFrame == MsoTriState.msoTrue)
            {
                var textFrame = shape.TextFrame;
                var textRange = textFrame.TextRange;
                var styleAttributes = new Dictionary<string, object>();
                
                // 텍스트 스타일 속성 가져오기
                styleAttributes["FontSize"] = textRange.Font.Size;
                styleAttributes["FontName"] = textRange.Font.Name;
                styleAttributes["FontWeight"] = textRange.Font.Bold == MsoTriState.msoTrue ? "Bold" : "Normal";
                styleAttributes["FontItalic"] = textRange.Font.Italic;
                styleAttributes["UnderlineStyle"] = textRange.Font.Underline;
                styleAttributes["ForegroundColor"] = textRange.Font.Color.RGB;
                
                // Shape의 배경색 사용
                if (shape.Fill.Visible == MsoTriState.msoTrue && shape.Fill.ForeColor.RGB != 0)
                {
                    styleAttributes["BackgroundColor"] = shape.Fill.ForeColor.RGB;
                    styleAttributes["BackgroundTransparency"] = shape.Fill.Transparency;
                }

                // TextFrame2를 사용하여 하이라이트 색상 가져오기 (Office 2019+)
                if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame2.HasText == MsoTriState.msoTrue)
                {
                    var tr2 = shape.TextFrame2.TextRange;
                    if (tr2.Font.Highlight.Type != MsoColorType.msoColorTypeMixed)
                    {
                        styleAttributes["HighlightColor"] = tr2.Font.Highlight.RGB;
                    }
                }
                
                // 텍스트 정렬 설정
                switch (textRange.ParagraphFormat.Alignment)
                {
                    case PpParagraphAlignment.ppAlignCenter:
                        styleAttributes["TextAlign"] = "center";
                        break;
                    case PpParagraphAlignment.ppAlignRight:
                        styleAttributes["TextAlign"] = "right";
                        break;
                    case PpParagraphAlignment.ppAlignJustify:
                        styleAttributes["TextAlign"] = "justify";
                        break;
                    default:
                        styleAttributes["TextAlign"] = "left";
                        break;
                }

                switch (textFrame.VerticalAnchor)
                {
                    case MsoVerticalAnchor.msoAnchorMiddle:
                        styleAttributes["VerticalAlign"] = "middle";
                        break;
                    case MsoVerticalAnchor.msoAnchorBottom:
                        styleAttributes["VerticalAlign"] = "bottom";
                        break;
                    default:
                        styleAttributes["VerticalAlign"] = "top";
                        break;
                }
                
                string textStyle = GetTextStyleString(styleAttributes);
                content = $"<div style='{textStyle}'>{textRange.Text}</div>";
            }
            else if (shape.Type == MsoShapeType.msoPicture)
            {
                content = $"<img src='data:image/png;base64,...' alt='Image' />";
            }

            return $"<{shapeType} style='{styleString}'>{content}</{shapeType}>";
        }

        public override (string SelectedText, Dictionary<string, object> StyleAttributes, string LineNumber) GetSelectedTextWithStyle(bool readAllContent = false)
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

                    Microsoft.Office.Interop.PowerPoint.ShapeRange? shapes;
                    if (readAllContent)
                    {
                        shapes = _slide.Shapes.Range();
                    }
                    else
                    {
                        shapes = _pptApp.ActiveWindow?.Selection?.ShapeRange;
                    }

                    if (shapes == null)
                    {
                        Console.WriteLine("선택된 텍스트가 없습니다.");
                        return (string.Empty, new Dictionary<string, object>(), string.Empty);
                    }

                    // 클립보드 형식 확인
                    var dataFormats = Forms.Clipboard.GetDataObject()?.GetFormats();
                    if (dataFormats != null)
                    {
                        Console.WriteLine("클립보드에 있는 데이터 형식:");
                        foreach (var format in dataFormats)
                        {
                            Console.WriteLine($"- {format}");
                        }
                    }

                    var styledTextBuilder = new StringBuilder();
                    var styleAttributes = new Dictionary<string, object>();

                    foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in shapes)
                    {
                        string shapeHtml = ConvertShapeToHtml(shape);
                        styledTextBuilder.Append(shapeHtml);
                    }

                    string selectedText = styledTextBuilder.ToString();
                    string lineNumber = $"Slide {_slide.SlideIndex}";

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

                        string fullHtml = string.Format(htmlTemplate, selectedText, selectedText.Length);
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
                        LogWindow.Instance.FilePathTextBox.Text = _presentation.FullName;
                        LogWindow.Instance.PositionTextBox.Text = lineNumber;
                        LogWindow.Instance.ContextTextBox.Text = selectedText;
                    });

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
