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
using HtmlAgilityPack;

namespace overlay_gpt
{
    public class PPTContextReader : BaseContextReader
    {
        private Application? _pptApp;
        private Presentation? _presentation;
        private Slide? _slide;
        private bool _isTargetProg;
        private string? _filePath;

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
            // HTML 태그가 포함된 텍스트를 안전하게 이스케이프 처리
            string result = System.Web.HttpUtility.HtmlEncode(text);
            var styles = new List<string>();
            var spanStyles = new List<string>();

            // 하이라이트 색상 처리
            if (styleAttributes.ContainsKey("HighlightColor"))
            {
                int highlightColor = Convert.ToInt32(styleAttributes["HighlightColor"]);
                if (highlightColor != 0)
                {
                    int rgbColor = ConvertColorToRGB(highlightColor);
                    int r = (rgbColor >> 16) & 0xFF;
                    int g = (rgbColor >> 8) & 0xFF;
                    int b = rgbColor & 0xFF;
                    spanStyles.Add($"background-color: rgb({r}, {g}, {b})");
                }
            }

            // 배경색 처리
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
                        spanStyles.Add($"background-color: rgba({r}, {g}, {b}, {1.0 - transparency})");
                    }
                }
            }

            // 기타 텍스트 스타일 처리
            if (styleAttributes.ContainsKey("UnderlineStyle"))
            {
                string underlineStyle = styleAttributes["UnderlineStyle"]?.ToString() ?? "none";
                if (underlineStyle == "underline")
                {
                    spanStyles.Add("text-decoration: underline");
                }
            }
            if (styleAttributes.ContainsKey("FontWeight") && styleAttributes["FontWeight"]?.ToString() == "Bold")
                spanStyles.Add("font-weight: bold");
            if (styleAttributes.ContainsKey("FontItalic") && styleAttributes["FontItalic"] is MsoTriState italic && italic == MsoTriState.msoTrue)
                spanStyles.Add("font-style: italic");
            if (styleAttributes.ContainsKey("FontStrikethrough") && styleAttributes["FontStrikethrough"] is MsoTriState strikethrough && strikethrough == MsoTriState.msoTrue)
                result = $"<s>{result}</s>";
            if (styleAttributes.ContainsKey("FontSize"))
            {
                double fontSize = Convert.ToDouble(styleAttributes["FontSize"]);
                spanStyles.Add($"font-size: {fontSize}pt");
            }

            // span 스타일이 있는 경우 span 태그로 감싸기
            if (spanStyles.Count > 0)
            {
                result = $"<span style='{string.Join("; ", spanStyles)}'>{result}</span>";
            }

            // div 스타일이 있는 경우 div 태그로 감싸기
            if (styles.Count > 0)
            {
                result = $"<div style='{string.Join("; ", styles)}'>{result}</div>";
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
            styleList.Add("width: 100%");
            styleList.Add("height: 100%");
            
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
            // 도형 타입 디버깅 로그 추가
            Console.WriteLine($"\n=== ConvertShapeToHtml 시작 ===");
            Console.WriteLine($"Shape Name: {shape.Name}");
            Console.WriteLine($"Shape Type: {shape.Type} ({shape.Type.ToString()})");
            Console.WriteLine($"Has Text Frame: {shape.HasTextFrame}");
            Console.WriteLine($"Shape Width: {shape.Width}, Height: {shape.Height}");
            Console.WriteLine($"Shape Left: {shape.Left}, Top: {shape.Top}");
            
            string shapeType = GetShapeType(shape);
            string styleString = GetShapeStyleString(shape);
            string content = string.Empty;
            
            Console.WriteLine($"Determined Shape Type: {shapeType}");
            Console.WriteLine($"Style String: {styleString}");

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
                
                // 밑줄 스타일 처리 - TextFrame2 사용
                if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame2.HasText == MsoTriState.msoTrue)
                {
                    var tr2 = shape.TextFrame2.TextRange;
                    if (tr2.Font.UnderlineStyle != MsoTextUnderlineType.msoNoUnderline)
                    {
                        switch (tr2.Font.UnderlineStyle)
                        {
                            case MsoTextUnderlineType.msoUnderlineSingleLine:
                            case MsoTextUnderlineType.msoUnderlineDoubleLine:
                            case MsoTextUnderlineType.msoUnderlineHeavyLine:
                            case MsoTextUnderlineType.msoUnderlineDottedLine:
                            case MsoTextUnderlineType.msoUnderlineDashLine:
                            case MsoTextUnderlineType.msoUnderlineDotDashLine:
                            case MsoTextUnderlineType.msoUnderlineDotDotDashLine:
                            case MsoTextUnderlineType.msoUnderlineWavyLine:
                            case MsoTextUnderlineType.msoUnderlineWavyHeavyLine:
                            case MsoTextUnderlineType.msoUnderlineWavyDoubleLine:
                                styleAttributes["UnderlineStyle"] = "underline";
                                break;
                            default:
                                styleAttributes["UnderlineStyle"] = "none";
                                break;
                        }
                    }
                    else
                    {
                        styleAttributes["UnderlineStyle"] = "none";
                    }

                    // 취소선 처리 - TextFrame2 사용
                    if (tr2.Font.StrikeThrough == MsoTriState.msoTrue)
                    {
                        styleAttributes["FontStrikethrough"] = MsoTriState.msoTrue;
                    }
                }
                else
                {
                    // 레거시 TextFrame 사용
                    styleAttributes["UnderlineStyle"] = textRange.Font.Underline == MsoTriState.msoTrue ? "underline" : "none";
                }
                
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
                content = GetStyledText(textRange.Text, styleAttributes);
                
                // 스타일 병합
                var mergedStyles = new Dictionary<string, string>();
                
                // 외부 div 스타일 파싱
                foreach (var style in styleString.Split(';', StringSplitOptions.RemoveEmptyEntries))
                {
                    var parts = style.Trim().Split(':');
                    if (parts.Length == 2)
                    {
                        mergedStyles[parts[0].Trim()] = parts[1].Trim();
                    }
                }
                
                // 내부 div 스타일 파싱 및 병합
                foreach (var style in textStyle.Split(';', StringSplitOptions.RemoveEmptyEntries))
                {
                    var parts = style.Trim().Split(':');
                    if (parts.Length == 2)
                    {
                        var key = parts[0].Trim();
                        var value = parts[1].Trim();
                        
                        // background-color와 font-size는 제외하고 다른 스타일만 병합
                        if (key != "background-color" && key != "font-size" && 
                            (key == "color" || key == "text-align" || key == "vertical-align"))
                        {
                            mergedStyles[key] = value;
                        }
                    }
                }
                
                // 병합된 스타일 문자열 생성
                string mergedStyleString = string.Join("; ", mergedStyles.Select(kv => $"{kv.Key}: {kv.Value}"));
                
                // 디버깅을 위한 출력
                Console.WriteLine("\n=== Shape HTML 변환 정보 ===");
                Console.WriteLine($"Shape Type: {shapeType}");
                Console.WriteLine($"Content: {content}");
                Console.WriteLine($"Merged Styles: {mergedStyleString}");
                Console.WriteLine($"Final HTML: <{shapeType} style='{mergedStyleString}'>{content}</{shapeType}>");
                Console.WriteLine("===========================\n");
                
                // 텍스트가 있는 도형에서도 배경 이미지 확인
                Console.WriteLine("=== 텍스트 도형에서 배경 이미지 확인 ===");
                try
                {
                    // Fill 속성에서 이미지 확인
                    if (shape.Fill.Visible == MsoTriState.msoTrue && 
                        shape.Fill.Type == MsoFillType.msoFillPicture)
                    {
                        Console.WriteLine("텍스트 도형에 배경 이미지 발견! 이미지 태그로 변환합니다.");
                        
                        // 이미지 저장 디렉토리 생성
                        string imageDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "images");
                        Console.WriteLine($"텍스트 도형 배경 이미지 디렉토리: {imageDir}");
                        if (!Directory.Exists(imageDir))
                        {
                            Directory.CreateDirectory(imageDir);
                            Console.WriteLine("텍스트 도형 배경 이미지 디렉토리 생성 완료");
                        }

                        // 임시 파일로 이미지 저장
                        string tempFile = Path.GetTempFileName() + ".png";
                        Console.WriteLine($"텍스트 도형 배경 이미지 임시 파일: {tempFile}");
                        
                        // 도형을 이미지로 Export (텍스트와 배경 이미지 모두 포함)
                        shape.Export(tempFile, PpShapeFormat.ppShapeFormatPNG);
                        Console.WriteLine("텍스트 도형 배경 이미지 Export 완료");
                        
                        // 이미지를 바이트 배열로 읽기
                        byte[] imageBytes = File.ReadAllBytes(tempFile);
                        Console.WriteLine($"텍스트 도형 배경 이미지 바이트 크기: {imageBytes.Length}");
                        
                        // 고유한 이미지 ID 생성
                        string imageId = Guid.NewGuid().ToString();
                        string imagePath = Path.Combine(imageDir, $"text_bg_{imageId}.png");
                        Console.WriteLine($"텍스트 도형 배경 이미지 최종 경로: {imagePath}");

                        // 이미지 데이터를 파일로 저장
                        File.WriteAllBytes(imagePath, imageBytes);
                        Console.WriteLine("텍스트 도형 배경 이미지 파일 저장 완료");
                        
                        // 임시 파일 삭제
                        File.Delete(tempFile);
                        Console.WriteLine("텍스트 도형 배경 이미지 임시 파일 삭제 완료");
                        
                        // 절대 경로로 이미지 참조
                        string absolutePath = Path.GetFullPath(imagePath);
                        Console.WriteLine($"텍스트 도형 배경 이미지 절대 경로: {absolutePath}");
                        string imgTag = $"<img style='{mergedStyleString}' src='{absolutePath}' alt='Text with Background Image' />";
                        Console.WriteLine($"생성된 텍스트 도형 배경 이미지 태그: {imgTag}");
                        Console.WriteLine("=== 텍스트 도형 배경 이미지 처리 완료 - 이미지 태그로 반환 ===");
                        return imgTag;
                    }
                    else
                    {
                        Console.WriteLine($"텍스트 도형에 배경 이미지 없음. Fill.Visible: {shape.Fill.Visible}, Fill.Type: {shape.Fill.Type}");
                        Console.WriteLine("일반 텍스트 도형으로 처리합니다.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"=== 텍스트 도형 배경 이미지 확인 중 오류 ===");
                    Console.WriteLine($"오류 메시지: {ex.Message}");
                    Console.WriteLine("일반 텍스트 도형으로 처리합니다.");
                }
                
                string resultHtml = $"<{shapeType} style='{mergedStyleString}'>{content}</{shapeType}>";
                Console.WriteLine($"=== 최종 결과 ===");
                Console.WriteLine($"Final HTML: {resultHtml}");
                Console.WriteLine($"=== ConvertShapeToHtml 완료 ===\n");
                return resultHtml;
            }
            else if (shape.Type == MsoShapeType.msoPicture)
            {
                Console.WriteLine("=== 이미지 처리 시작 ===");
                try
                {
                    // 이미지 저장 디렉토리 생성
                    string imageDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "images");
                    Console.WriteLine($"이미지 디렉토리: {imageDir}");
                    if (!Directory.Exists(imageDir))
                    {
                        Directory.CreateDirectory(imageDir);
                        Console.WriteLine("이미지 디렉토리 생성 완료");
                    }

                    // 임시 파일로 이미지 저장
                    string tempFile = Path.GetTempFileName() + ".png";
                    Console.WriteLine($"임시 파일 경로: {tempFile}");
                    
                    shape.Export(tempFile, PpShapeFormat.ppShapeFormatPNG);
                    Console.WriteLine("이미지 Export 완료");
                    
                    // 이미지를 바이트 배열로 읽기
                    byte[] imageBytes = File.ReadAllBytes(tempFile);
                    Console.WriteLine($"이미지 바이트 크기: {imageBytes.Length}");
                    
                    // 고유한 이미지 ID 생성
                    string imageId = Guid.NewGuid().ToString();
                    string imagePath = Path.Combine(imageDir, $"{imageId}.png");
                    Console.WriteLine($"최종 이미지 경로: {imagePath}");

                    // 이미지 데이터를 파일로 저장
                    File.WriteAllBytes(imagePath, imageBytes);
                    Console.WriteLine("이미지 파일 저장 완료");
                    
                    // 임시 파일 삭제
                    File.Delete(tempFile);
                    Console.WriteLine("임시 파일 삭제 완료");
                    
                    // 절대 경로로 이미지 참조
                    string absolutePath = Path.GetFullPath(imagePath);
                    Console.WriteLine($"절대 경로: {absolutePath}");
                    string imgTag = $"<img style='{styleString}' src='{absolutePath}' alt='Image' />";
                    Console.WriteLine($"생성된 img 태그: {imgTag}");
                    Console.WriteLine("=== 이미지 처리 완료 ===");
                    return imgTag;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"=== 이미지 변환 중 오류 발생 ===");
                    Console.WriteLine($"오류 메시지: {ex.Message}");
                    Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                    string errorImgTag = $"<img style='{styleString}' src='' alt='Image' />";
                    Console.WriteLine($"오류 시 반환할 태그: {errorImgTag}");
                    Console.WriteLine("=== 이미지 처리 오류 완료 ===");
                    return errorImgTag;
                }
            }
            else
            {
                // 다른 도형 타입에서도 배경 이미지 확인
                Console.WriteLine("=== 일반 도형에서 배경 이미지 확인 ===");
                try
                {
                    // Fill 속성에서 이미지 확인
                    if (shape.Fill.Visible == MsoTriState.msoTrue && 
                        shape.Fill.Type == MsoFillType.msoFillPicture)
                    {
                        Console.WriteLine("도형에 배경 이미지 발견!");
                        
                        // 이미지 저장 디렉토리 생성
                        string imageDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "images");
                        Console.WriteLine($"배경 이미지 디렉토리: {imageDir}");
                        if (!Directory.Exists(imageDir))
                        {
                            Directory.CreateDirectory(imageDir);
                            Console.WriteLine("배경 이미지 디렉토리 생성 완료");
                        }

                        // 임시 파일로 이미지 저장
                        string tempFile = Path.GetTempFileName() + ".png";
                        Console.WriteLine($"배경 이미지 임시 파일: {tempFile}");
                        
                        // 도형을 이미지로 Export (배경 이미지 포함)
                        shape.Export(tempFile, PpShapeFormat.ppShapeFormatPNG);
                        Console.WriteLine("배경 이미지 Export 완료");
                        
                        // 이미지를 바이트 배열로 읽기
                        byte[] imageBytes = File.ReadAllBytes(tempFile);
                        Console.WriteLine($"배경 이미지 바이트 크기: {imageBytes.Length}");
                        
                        // 고유한 이미지 ID 생성
                        string imageId = Guid.NewGuid().ToString();
                        string imagePath = Path.Combine(imageDir, $"bg_{imageId}.png");
                        Console.WriteLine($"배경 이미지 최종 경로: {imagePath}");

                        // 이미지 데이터를 파일로 저장
                        File.WriteAllBytes(imagePath, imageBytes);
                        Console.WriteLine("배경 이미지 파일 저장 완료");
                        
                        // 임시 파일 삭제
                        File.Delete(tempFile);
                        Console.WriteLine("배경 이미지 임시 파일 삭제 완료");
                        
                        // 절대 경로로 이미지 참조
                        string absolutePath = Path.GetFullPath(imagePath);
                        Console.WriteLine($"배경 이미지 절대 경로: {absolutePath}");
                        string imgTag = $"<img style='{styleString}' src='{absolutePath}' alt='Background Image' />";
                        Console.WriteLine($"생성된 배경 이미지 태그: {imgTag}");
                        Console.WriteLine("=== 배경 이미지 처리 완료 ===");
                        return imgTag;
                    }
                    else
                    {
                        Console.WriteLine($"배경 이미지 없음. Fill.Visible: {shape.Fill.Visible}, Fill.Type: {shape.Fill.Type}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"=== 배경 이미지 확인 중 오류 ===");
                    Console.WriteLine($"오류 메시지: {ex.Message}");
                    Console.WriteLine("=== 배경 이미지 확인 오류 완료 ===");
                }
                
                // msoPlaceholder 타입의 이미지 placeholder 확인
                Console.WriteLine("=== Placeholder 이미지 확인 ===");
                try
                {
                    if (shape.Type == MsoShapeType.msoPlaceholder)
                    {
                        Console.WriteLine($"Placeholder 도형 발견. Name: {shape.Name}");
                        
                        // 이름에 "Picture"가 포함되어 있거나 PlaceholderFormat이 이미지 타입인지 확인
                        bool isImagePlaceholder = false;
                        
                        // 이름 확인
                        if (shape.Name.ToLower().Contains("picture"))
                        {
                            Console.WriteLine("이름에 'Picture'가 포함된 placeholder 발견");
                            isImagePlaceholder = true;
                        }
                        
                        // PlaceholderFormat 확인
                        try
                        {
                            if (shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderPicture ||
                                shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderBitmap)
                            {
                                Console.WriteLine($"이미지 타입 Placeholder 발견: {shape.PlaceholderFormat.Type}");
                                isImagePlaceholder = true;
                            }
                        }
                        catch (Exception placeholderEx)
                        {
                            Console.WriteLine($"PlaceholderFormat 확인 중 오류 (무시): {placeholderEx.Message}");
                        }
                        
                        if (isImagePlaceholder)
                        {
                            Console.WriteLine("이미지 Placeholder로 판단됨. 도형을 이미지로 Export합니다.");
                            
                            // 이미지 저장 디렉토리 생성
                            string imageDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "images");
                            Console.WriteLine($"Placeholder 이미지 디렉토리: {imageDir}");
                            if (!Directory.Exists(imageDir))
                            {
                                Directory.CreateDirectory(imageDir);
                                Console.WriteLine("Placeholder 이미지 디렉토리 생성 완료");
                            }

                            // 임시 파일로 이미지 저장
                            string tempFile = Path.GetTempFileName() + ".png";
                            Console.WriteLine($"Placeholder 이미지 임시 파일: {tempFile}");
                            
                            // 도형을 이미지로 Export
                            shape.Export(tempFile, PpShapeFormat.ppShapeFormatPNG);
                            Console.WriteLine("Placeholder 이미지 Export 완료");
                            
                            // 이미지를 바이트 배열로 읽기
                            byte[] imageBytes = File.ReadAllBytes(tempFile);
                            Console.WriteLine($"Placeholder 이미지 바이트 크기: {imageBytes.Length}");
                            
                            // 고유한 이미지 ID 생성
                            string imageId = Guid.NewGuid().ToString();
                            string imagePath = Path.Combine(imageDir, $"placeholder_{imageId}.png");
                            Console.WriteLine($"Placeholder 이미지 최종 경로: {imagePath}");

                            // 이미지 데이터를 파일로 저장
                            File.WriteAllBytes(imagePath, imageBytes);
                            Console.WriteLine("Placeholder 이미지 파일 저장 완료");
                            
                            // 임시 파일 삭제
                            File.Delete(tempFile);
                            Console.WriteLine("Placeholder 이미지 임시 파일 삭제 완료");
                            
                            // 절대 경로로 이미지 참조
                            string absolutePath = Path.GetFullPath(imagePath);
                            Console.WriteLine($"Placeholder 이미지 절대 경로: {absolutePath}");
                            string imgTag = $"<img style='{styleString}' src='{absolutePath}' alt='Placeholder Image' />";
                            Console.WriteLine($"생성된 Placeholder 이미지 태그: {imgTag}");
                            Console.WriteLine("=== Placeholder 이미지 처리 완료 ===");
                            return imgTag;
                        }
                        else
                        {
                            Console.WriteLine("이미지 Placeholder가 아님");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"=== Placeholder 이미지 확인 중 오류 ===");
                    Console.WriteLine($"오류 메시지: {ex.Message}");
                    Console.WriteLine("=== Placeholder 이미지 확인 오류 완료 ===");
                }
                
                // 그룹 도형인 경우 내부 도형들 확인
                if (shape.Type == MsoShapeType.msoGroup)
                {
                    Console.WriteLine("=== 그룹 도형에서 이미지 검색 ===");
                    try
                    {
                        foreach (Microsoft.Office.Interop.PowerPoint.Shape groupShape in shape.GroupItems)
                        {
                            Console.WriteLine($"그룹 내부 도형 타입: {groupShape.Type}");
                            
                            if (groupShape.Type == MsoShapeType.msoPicture)
                            {
                                Console.WriteLine("그룹 내부에서 이미지 발견!");
                                
                                // 이미지 저장 디렉토리 생성
                                string imageDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "images");
                                if (!Directory.Exists(imageDir))
                                {
                                    Directory.CreateDirectory(imageDir);
                                }

                                // 임시 파일로 이미지 저장
                                string tempFile = Path.GetTempFileName() + ".png";
                                
                                // 그룹 내부 이미지 Export
                                groupShape.Export(tempFile, PpShapeFormat.ppShapeFormatPNG);
                                Console.WriteLine("그룹 내부 이미지 Export 완료");
                                
                                // 이미지를 바이트 배열로 읽기
                                byte[] imageBytes = File.ReadAllBytes(tempFile);
                                Console.WriteLine($"그룹 내부 이미지 바이트 크기: {imageBytes.Length}");
                                
                                // 고유한 이미지 ID 생성
                                string imageId = Guid.NewGuid().ToString();
                                string imagePath = Path.Combine(imageDir, $"group_{imageId}.png");

                                // 이미지 데이터를 파일로 저장
                                File.WriteAllBytes(imagePath, imageBytes);
                                Console.WriteLine("그룹 내부 이미지 파일 저장 완료");
                                
                                // 임시 파일 삭제
                                File.Delete(tempFile);
                                
                                // 절대 경로로 이미지 참조
                                string absolutePath = Path.GetFullPath(imagePath);
                                string imgTag = $"<img style='{styleString}' src='{absolutePath}' alt='Group Image' />";
                                Console.WriteLine($"그룹 내부 이미지 태그: {imgTag}");
                                Console.WriteLine("=== 그룹 내부 이미지 처리 완료 ===");
                                return imgTag;
                            }
                        }
                        Console.WriteLine("그룹 내부에 이미지 없음");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"=== 그룹 도형 처리 중 오류 ===");
                        Console.WriteLine($"오류 메시지: {ex.Message}");
                        Console.WriteLine("=== 그룹 도형 처리 오류 완료 ===");
                    }
                }
            }

            string finalHtml = $"<{shapeType} style='{styleString}'>{content}</{shapeType}>";
            Console.WriteLine($"=== 최종 결과 ===");
            Console.WriteLine($"Final HTML: {finalHtml}");
            Console.WriteLine($"=== ConvertShapeToHtml 완료 ===\n");
            return finalHtml;
        }

        private bool IsPowerPointProcessActive()
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

                // PowerPoint 프로세스 이름 확인 (POWERPNT.EXE)
                return foregroundProcess.ProcessName.Equals("POWERPNT", StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"프로세스 확인 중 오류 발생: {ex.Message}");
                return false;
            }
        }

        public PPTContextReader(bool isTargetProg = false, string filePath = "")
        {
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            Console.WriteLine($"[{timestamp}] === PPTContextReader 생성자 호출 ===");
            Console.WriteLine($"[{timestamp}] PPTContextReader 생성 시도 - isTargetProg: {isTargetProg}");
            Console.WriteLine($"[{timestamp}] 파일 경로: {filePath}");
            Console.WriteLine($"[{timestamp}] 스레드 ID: {System.Threading.Thread.CurrentThread.ManagedThreadId}");
            Console.WriteLine($"[{timestamp}] 프로세스 ID: {System.Diagnostics.Process.GetCurrentProcess().Id}");
            _isTargetProg = isTargetProg;
            _filePath = filePath;
            Console.WriteLine($"[{timestamp}] === PPTContextReader 생성자 완료 ===\n");
        }

        public override (string SelectedText, Dictionary<string, object> StyleAttributes, string LineNumber) GetSelectedTextWithStyle(bool readAllContent = false)
        {
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            Console.WriteLine($"[{timestamp}] === GetSelectedTextWithStyle 시작 ===");
            Console.WriteLine($"[{timestamp}] readAllContent: {readAllContent}");
            Console.WriteLine($"[{timestamp}] isTargetProg: {_isTargetProg}");
            Console.WriteLine($"[{timestamp}] 스레드 ID: {System.Threading.Thread.CurrentThread.ManagedThreadId}");
            Console.WriteLine($"[{timestamp}] 프로세스 ID: {System.Diagnostics.Process.GetCurrentProcess().Id}");
            
            var styledTextBuilder = new StringBuilder();
            var styleAttributes = new Dictionary<string, object>();
            string selectedText = string.Empty;
            string lineNumber = string.Empty;

            try
            {
                Console.WriteLine("PowerPoint 데이터 읽기 시작...");
                Console.WriteLine($"readAllContent: {readAllContent}");
                Console.WriteLine($"isTargetProg: {_isTargetProg}");

                // isTargetProg가 false일 때만 현재 포커스된 프로세스가 PowerPoint인지 확인
                if (!_isTargetProg && !IsPowerPointProcessActive())
                {
                    Console.WriteLine("현재 포커스된 프로세스가 PowerPoint가 아닙니다.");
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);
                }

                var pptProcesses = Process.GetProcessesByName("POWERPNT");
                Console.WriteLine($"실행 중인 PowerPoint 프로세스 수: {pptProcesses.Length}");
                
                try
                {
                    Console.WriteLine("PowerPoint COM 객체 가져오기 시도...");
                    _pptApp = (Application)GetActiveObject("PowerPoint.Application");
                    Console.WriteLine("PowerPoint COM 객체 가져오기 성공");

                    if (_isTargetProg && !string.IsNullOrEmpty(_filePath))
                    {
                        // 기존 프로세스에서 원하는 파일 찾기
                        bool found = false;
                        foreach (Presentation pres in _pptApp.Presentations)
                        {
                            if (string.Equals(pres.FullName, _filePath, StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine($"기존 프로세스에서 파일 찾음: {_filePath}");
                                _presentation = pres;
                                found = true;
                                break;
                            }
                        }

                        // 파일을 찾지 못했다면 새로 열기
                        if (!found)
                        {
                            Console.WriteLine($"기존 프로세스에서 파일을 찾지 못해 새로 열기 시도: {_filePath}");
                            _presentation = _pptApp.Presentations.Open(_filePath);
                            Console.WriteLine("파일 열기 성공");
                        }
                    }
                    else
                    {
                        _presentation = _pptApp.ActivePresentation;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"PowerPoint COM 객체 가져오기 실패: {ex.Message}");
                    if (_isTargetProg)
                    {
                        Console.WriteLine("새 PowerPoint 프로세스를 생성합니다.");
                        try
                        {
                            _pptApp = new Application();
                            Console.WriteLine("새 PowerPoint 애플리케이션 생성 성공");

                            if (!string.IsNullOrEmpty(_filePath))
                            {
                                Console.WriteLine($"파일 열기 시도: {_filePath}");
                                _presentation = _pptApp.Presentations.Open(_filePath);
                                Console.WriteLine("파일 열기 성공");
                            }
                        }
                        catch (Exception createEx)
                        {
                            Console.WriteLine($"새 PowerPoint 애플리케이션 생성 실패: {createEx.Message}");
                            Console.WriteLine($"스택 트레이스: {createEx.StackTrace}");
                            throw;
                        }
                    }
                    else
                    {
                        throw new InvalidOperationException("PowerPoint is not running");
                    }
                }

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

                if(_isTargetProg)
                {
                    Console.WriteLine("전체 슬라이드 선택");
                    try 
                    {
                        // PowerPoint를 일시적으로 보이게 설정
                        bool originalVisible = _pptApp.Visible == MsoTriState.msoTrue;
                        _pptApp.Visible = MsoTriState.msoTrue;
                        
                        // 전체 슬라이드 순회
                        foreach (Slide slide in _presentation.Slides)
                        {
                            _slide = slide;
                            var allShapes = _slide.Shapes.Range();
                            Console.WriteLine($"슬라이드 {_slide.SlideIndex} 선택 완료");
                            
                            // 슬라이드 시작 div 추가
                            styledTextBuilder.Append($"<div class='Slide{_slide.SlideIndex}'>");
                            
                            foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in allShapes)
                            {
                                string shapeHtml = ConvertShapeToHtml(shape);
                                styledTextBuilder.Append(shapeHtml);
                            }
                            
                            // 슬라이드 종료 div
                            styledTextBuilder.Append("</div>");
                        }

                        selectedText = styledTextBuilder.ToString();
                        lineNumber = $"All Slides ({_presentation.Slides.Count} slides)";

                        // 원래 상태로 복원
                        _pptApp.Visible = originalVisible ? MsoTriState.msoTrue : MsoTriState.msoFalse;

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

                        return (selectedText, styleAttributes, lineNumber);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"전체 슬라이드 선택 중 오류 발생: {ex.Message}");
                        Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                        return (string.Empty, new Dictionary<string, object>(), string.Empty);
                    }
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

                foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in shapes)
                {
                    string shapeHtml = ConvertShapeToHtml(shape);
                    styledTextBuilder.Append(shapeHtml);
                }

                selectedText = styledTextBuilder.ToString();
                // HTML 단순화 적용
                lineNumber = $"Slide {_slide.SlideIndex}";

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
            Application? tempPptApp = null;
            Presentation? tempPresentation = null;
            
            try
            {
                Console.WriteLine("PowerPoint COM 객체 가져오기 시도...");
                tempPptApp = (Application)GetActiveObject("PowerPoint.Application");
                Console.WriteLine("PowerPoint COM 객체 가져오기 성공");

                if (_isTargetProg && !string.IsNullOrEmpty(_filePath))
                {
                    // 모든 프레젠테이션 확인
                    foreach (Presentation pres in tempPptApp.Presentations)
                    {
                        try
                        {
                            string filePath = pres.FullName;
                            string fileName = pres.Name;
                            
                            Console.WriteLine($"PowerPoint 문서 정보:");
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
                                    "PowerPoint",
                                    fileName,
                                    filePath
                                );
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"프레젠테이션 처리 중 오류: {ex.Message}");
                            continue;
                        }
                    }
                }
                else
                {
                    // 활성 프레젠테이션 정보 가져오기
                    tempPresentation = tempPptApp.ActivePresentation;
                    if (tempPresentation != null)
                    {
                        string filePath = tempPresentation.FullName;
                        string fileName = tempPresentation.Name;
                        
                        Console.WriteLine($"활성 PowerPoint 문서 정보:");
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
                            "PowerPoint",
                            fileName,
                            filePath
                        );
                    }
                }
                
                Console.WriteLine("문서를 찾을 수 없습니다.");
                return (null, null, "PowerPoint", string.Empty, string.Empty);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"파일 정보 가져오기 오류: {ex.Message}");
                Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                return (null, null, "PowerPoint", string.Empty, string.Empty);
            }
            finally
            {
                if (tempPresentation != null) Marshal.ReleaseComObject(tempPresentation);
                if (tempPptApp != null) Marshal.ReleaseComObject(tempPptApp);
            }
        }
    }
}
