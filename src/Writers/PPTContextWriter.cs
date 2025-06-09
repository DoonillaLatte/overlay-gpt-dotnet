using System;
using System.Collections.Generic;
using System.Windows.Automation;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.IO;
using Forms = System.Windows.Forms;
using HtmlAgilityPack;

namespace overlay_gpt
{
    public class PPTContextWriter : IContextWriter
    {
        private Application? _pptApp;
        private Presentation? _presentation;
        private Slide? _slide;
        private bool _isTargetProg;

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("oleaut32.dll")]
        private static extern int GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        [DllImport("ole32.dll")]
        private static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);

        private static object GetActiveObject(string progID)
        {
            Guid clsid;
            CLSIDFromProgID(progID, out clsid);
            object obj;
            GetActiveObject(ref clsid, IntPtr.Zero, out obj);
            return obj;
        }

        public bool IsTargetProg
        {
            get => _isTargetProg;
            set => _isTargetProg = value;
        }

        public bool OpenFile(string filePath)
        {
            try
            {
                Console.WriteLine($"PowerPoint 파일 열기 시도: {filePath}");

                if (!File.Exists(filePath))
                {
                    Console.WriteLine("파일이 존재하지 않습니다.");
                    return false;
                }

                // 기존 PowerPoint 프로세스 확인
                try
                {
                    _pptApp = (Application)GetActiveObject("PowerPoint.Application");
                    
                    // 이미 열려있는 프레젠테이션 확인
                    foreach (Presentation pres in _pptApp.Presentations)
                    {
                        if (pres.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase))
                        {
                            _presentation = pres;
                            _slide = _pptApp.ActiveWindow?.View?.Slide;
                            Console.WriteLine("이미 열려있는 PowerPoint 파일을 사용합니다.");
                            return true;
                        }
                    }
                }
                catch
                {
                    // PowerPoint가 실행중이 아닌 경우 새로 시작
                    _pptApp = new Application();
                }

                // 새로 파일 열기
                _presentation = _pptApp.Presentations.Open(filePath, WithWindow: MsoTriState.msoTrue);
                _slide = _pptApp.ActiveWindow?.View?.Slide;

                Console.WriteLine("PowerPoint 파일 열기 성공");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"PowerPoint 파일 열기 오류: {ex.Message}");
                return false;
            }
        }

        private float ParseSafeFloat(string value, float defaultValue = 0)
        {
            try
            {
                var cleanValue = value?.Replace("px", "").Replace("pt", "").Trim();
                if (string.IsNullOrEmpty(cleanValue))
                {
                    Console.WriteLine($"빈 값이 전달됨, 기본값 {defaultValue} 사용");
                    return defaultValue;
                }
                    
                if (float.TryParse(cleanValue, out float result))
                {
                    Console.WriteLine($"값 파싱 성공: '{value}' -> {result}");
                    return result;
                }
                Console.WriteLine($"값 파싱 실패: '{value}', 기본값 {defaultValue} 사용");
                return defaultValue;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"값 파싱 중 예외: '{value}' - {ex.Message}, 기본값 {defaultValue} 사용");
                return defaultValue;
            }
        }

        private bool IsValidPosition(float value)
        {
            bool isValid = value >= -10000 && value <= 10000;
            if (!isValid)
                Console.WriteLine($"유효하지 않은 위치 값: {value} (유효 범위: -10000 ~ 10000)");
            return isValid;
        }

        private bool IsValidSize(float value)
        {
            bool isValid = value >= 1 && value <= 5000;
            if (!isValid)
                Console.WriteLine($"유효하지 않은 크기 값: {value} (유효 범위: 1 ~ 5000)");
            return isValid;
        }

        private void ApplyStyleToShape(Microsoft.Office.Interop.PowerPoint.Shape shape, string style)
        {
            var styleDict = new Dictionary<string, string>();
            foreach (var stylePart in style.Split(';'))
            {
                var parts = stylePart.Split(':');
                if (parts.Length == 2)
                {
                    styleDict[parts[0].Trim()] = parts[1].Trim();
                }
            }

            // 위치와 크기 적용 (안전한 파싱과 검증)
            try
            {
                if (styleDict.ContainsKey("left"))
                {
                    var leftValue = ParseSafeFloat(styleDict["left"]);
                    if (IsValidPosition(leftValue))
                        shape.Left = leftValue;
                }
                if (styleDict.ContainsKey("top"))
                {
                    var topValue = ParseSafeFloat(styleDict["top"]);
                    if (IsValidPosition(topValue))
                        shape.Top = topValue;
                }
                if (styleDict.ContainsKey("width"))
                {
                    var widthValue = ParseSafeFloat(styleDict["width"], 100);
                    if (IsValidSize(widthValue))
                        shape.Width = widthValue;
                }
                if (styleDict.ContainsKey("height"))
                {
                    var heightValue = ParseSafeFloat(styleDict["height"], 50);
                    if (IsValidSize(heightValue))
                        shape.Height = heightValue;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"위치/크기 적용 중 오류: {ex.Message}");
            }

            // 회전 적용 (안전한 파싱)
            try
            {
                if (styleDict.ContainsKey("transform"))
                {
                    var transform = styleDict["transform"];
                    if (transform.Contains("rotate"))
                    {
                        var rotation = transform.Replace("rotate(", "").Replace("deg)", "");
                        var rotationValue = ParseSafeFloat(rotation);
                        if (rotationValue >= -360 && rotationValue <= 360)
                            shape.Rotation = rotationValue;
                    }
                    else if (transform.Contains("rotateX"))
                    {
                        var rotationX = transform.Replace("rotateX(", "").Replace("deg)", "");
                        var rotationXValue = ParseSafeFloat(rotationX);
                        if (rotationXValue >= -360 && rotationXValue <= 360)
                            shape.ThreeD.RotationX = rotationXValue;
                    }
                    else if (transform.Contains("rotateY"))
                    {
                        var rotationY = transform.Replace("rotateY(", "").Replace("deg)", "");
                        var rotationYValue = ParseSafeFloat(rotationY);
                        if (rotationYValue >= -360 && rotationYValue <= 360)
                            shape.ThreeD.RotationY = rotationYValue;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"회전 적용 중 오류: {ex.Message}");
            }

            // 3D 효과 적용 (안전한 파싱)
            try
            {
                if (styleDict.ContainsKey("transform-style") && styleDict["transform-style"] == "preserve-3d")
                {
                    shape.ThreeD.Visible = MsoTriState.msoTrue;
                }
                if (styleDict.ContainsKey("perspective"))
                {
                    var perspectiveValue = ParseSafeFloat(styleDict["perspective"]);
                    if (perspectiveValue >= 0 && perspectiveValue <= 5000)
                    {
                        shape.ThreeD.Perspective = (MsoTriState)perspectiveValue;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"3D 효과 적용 중 오류: {ex.Message}");
            }

            // 텍스트 스타일 적용
            if (styleDict.ContainsKey("font-family"))
            {
                Console.WriteLine($"폰트 적용: {styleDict["font-family"]}");
                shape.TextFrame.TextRange.Font.Name = styleDict["font-family"];
            }
            if (styleDict.ContainsKey("font-weight"))
            {
                Console.WriteLine($"굵게 적용: {styleDict["font-weight"]}");
                shape.TextFrame.TextRange.Font.Bold = styleDict["font-weight"] == "bold" ? MsoTriState.msoTrue : MsoTriState.msoFalse;
            }
            if (styleDict.ContainsKey("font-style"))
            {
                Console.WriteLine($"기울임 적용: {styleDict["font-style"]}");
                shape.TextFrame.TextRange.Font.Italic = styleDict["font-style"] == "italic" ? MsoTriState.msoTrue : MsoTriState.msoFalse;
            }
            if (styleDict.ContainsKey("text-decoration"))
            {
                var decoration = styleDict["text-decoration"];
                Console.WriteLine($"텍스트 장식 적용: {decoration}");
                if (decoration.Contains("underline"))
                {
                    Console.WriteLine("밑줄 적용");
                    if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame2.HasText == MsoTriState.msoTrue)
                    {
                        shape.TextFrame2.TextRange.Font.UnderlineStyle = MsoTextUnderlineType.msoUnderlineSingleLine;
                    }
                    else
                    {
                        shape.TextFrame.TextRange.Font.Underline = MsoTriState.msoTrue;
                    }
                }
                if (decoration.Contains("line-through"))
                {
                    Console.WriteLine("취소선 적용");
                    if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame2.HasText == MsoTriState.msoTrue)
                    {
                        shape.TextFrame2.TextRange.Font.StrikeThrough = MsoTriState.msoTrue;
                    }
                }
            }

            // 텍스트 색상 적용
            if (styleDict.ContainsKey("color"))
            {
                var color = styleDict["color"];
                if (color.StartsWith("#"))
                {
                    shape.TextFrame.TextRange.Font.Color.RGB = ParseColor(color);
                }
            }
            else
            {
                shape.TextFrame.TextRange.Font.Color.RGB = 0;
            }

            // 폰트 크기 적용 (안전한 파싱)
            try
            {
                if (styleDict.ContainsKey("font-size"))
                {
                    var fontSize = ParseSafeFloat(styleDict["font-size"], 11);
                    if (fontSize >= 1 && fontSize <= 1638)
                    {
                        shape.TextFrame.TextRange.Font.Size = fontSize;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"폰트 크기 적용 중 오류: {ex.Message}");
            }

            // 배경색 적용
            if (styleDict.ContainsKey("background-color"))
            {
                var bgColor = styleDict["background-color"];
                if (bgColor.StartsWith("rgba"))
                {
                    var rgba = bgColor.Replace("rgba(", "").Replace(")", "").Split(',');
                    if (rgba.Length == 4)
                    {
                        int r = int.Parse(rgba[0].Trim());
                        int g = int.Parse(rgba[1].Trim());
                        int b = int.Parse(rgba[2].Trim());
                        float a = float.Parse(rgba[3].Trim());
                        
                        if (a < 0.1f)
                        {
                            shape.Fill.Visible = MsoTriState.msoFalse;
                        }
                        else
                        {
                            shape.Fill.Visible = MsoTriState.msoTrue;
                            shape.Fill.ForeColor.RGB = (b << 16) | (g << 8) | r;
                            shape.Fill.Transparency = 1 - a;
                        }
                    }
                }
                else if (bgColor.StartsWith("rgb"))
                {
                    var rgb = bgColor.Replace("rgb(", "").Replace(")", "").Split(',');
                    if (rgb.Length == 3)
                    {
                        int r = int.Parse(rgb[0].Trim());
                        int g = int.Parse(rgb[1].Trim());
                        int b = int.Parse(rgb[2].Trim());
                        shape.Fill.Visible = MsoTriState.msoTrue;
                        shape.Fill.ForeColor.RGB = (b << 16) | (g << 8) | r;
                        shape.Fill.Transparency = 0;
                    }
                }
                else if (bgColor.StartsWith("linear-gradient"))
                {
                    // 그라데이션 적용
                    var gradient = bgColor.Replace("linear-gradient(", "").Replace(")", "");
                    var parts = gradient.Split(',');
                    if (parts.Length >= 2)
                    {
                        shape.Fill.GradientStops.Insert(ParseColor(parts[1].Trim()), 0);
                        shape.Fill.GradientStops.Insert(ParseColor(parts[2].Trim()), 1);
                    }
                }
            }
            else
            {
                shape.Fill.Visible = MsoTriState.msoFalse;
            }

            // 테두리 적용 (안전한 파싱)
            try
            {
                if (styleDict.ContainsKey("border"))
                {
                    var border = styleDict["border"].Split(' ');
                    if (border.Length >= 3)
                    {
                        var borderWeight = ParseSafeFloat(border[0], 1);
                        if (borderWeight >= 0 && borderWeight <= 100)
                        {
                            shape.Line.Weight = borderWeight;
                            shape.Line.ForeColor.RGB = ParseColor(border[2]);
                        }
                    }
                }
                else
                {
                    shape.Line.Visible = MsoTriState.msoFalse;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"테두리 적용 중 오류: {ex.Message}");
                shape.Line.Visible = MsoTriState.msoFalse;
            }

            // 모서리 둥글기 적용 (안전한 파싱)
            try
            {
                if (styleDict.ContainsKey("border-radius"))
                {
                    var radius = ParseSafeFloat(styleDict["border-radius"]);
                    if (radius >= 0 && radius <= 100)
                    {
                        shape.Adjustments[1] = radius;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"모서리 둥글기 적용 중 오류: {ex.Message}");
            }

            // 그림자 효과 적용 (안전한 파싱)
            try
            {
                if (styleDict.ContainsKey("box-shadow"))
                {
                    var shadow = styleDict["box-shadow"].Split(' ');
                    if (shadow.Length >= 4)
                    {
                        var offsetX = ParseSafeFloat(shadow[0]);
                        var offsetY = ParseSafeFloat(shadow[1]);
                        var blur = ParseSafeFloat(shadow[2]);
                        
                        if (offsetX >= -100 && offsetX <= 100 && 
                            offsetY >= -100 && offsetY <= 100 && 
                            blur >= 0 && blur <= 100)
                        {
                            shape.Shadow.Visible = MsoTriState.msoTrue;
                            shape.Shadow.OffsetX = offsetX;
                            shape.Shadow.OffsetY = offsetY;
                            shape.Shadow.Blur = blur;
                            shape.Shadow.ForeColor.RGB = ParseColor(shadow[3]);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"그림자 효과 적용 중 오류: {ex.Message}");
            }

            // Z-인덱스 적용 (안전한 파싱)
            try
            {
                if (styleDict.ContainsKey("z-index"))
                {
                    var zIndex = (int)ParseSafeFloat(styleDict["z-index"]);
                    if (zIndex >= 0 && zIndex <= 100)
                    {
                        // Z-인덱스는 읽기 전용이므로 ZOrder 메서드를 사용
                        shape.ZOrder(MsoZOrderCmd.msoBringToFront);
                        for (int i = 0; i < zIndex; i++)
                        {
                            shape.ZOrder(MsoZOrderCmd.msoSendBackward);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Z-인덱스 적용 중 오류: {ex.Message}");
            }
        }

        private int ParseColor(string color)
        {
            if (color.StartsWith("#"))
            {
                color = color.Substring(1);
                int rgb = Convert.ToInt32(color, 16);
                // RGB 순서를 BGR로 변경
                int r = (rgb >> 16) & 0xFF;
                int g = (rgb >> 8) & 0xFF;
                int b = rgb & 0xFF;
                return (b << 16) | (g << 8) | r;
            }
            return 0;
        }

        public bool ApplyTextWithStyle(string text, string lineNumber)
        {
            try
            {
                Console.WriteLine("HTML 텍스트 적용 시작...");

                if (_presentation == null)
                {
                    Console.WriteLine("PowerPoint가 열려있지 않습니다.");
                    return false;
                }

                // 전체 슬라이드 처리
                if (lineNumber.StartsWith("All Slides"))
                {
                    var doc = new HtmlDocument();
                    doc.LoadHtml(text);

                    // 각 슬라이드별 div 찾기
                    for (int i = 1; i <= _presentation.Slides.Count; i++)
                    {
                        var slideDiv = doc.DocumentNode.SelectSingleNode($"//div[contains(@class, 'Slide{i}')]");
                        if (slideDiv != null)
                        {
                            _slide = _presentation.Slides[i];
                            
                            // 기존 도형 삭제
                            while (_slide.Shapes.Count > 0)
                            {
                                _slide.Shapes[1].Delete();
                            }

                            // 슬라이드 내용 적용
                            foreach (var node in slideDiv.ChildNodes)
                            {
                                if (node.NodeType == HtmlNodeType.Element)
                                {
                                    ProcessHtmlNode(node);
                                }
                            }
                        }
                    }

                    // 추가 슬라이드가 필요한 경우 새로 생성
                    var maxSlideNumber = 0;
                    foreach (var node in doc.DocumentNode.SelectNodes("//div[contains(@class, 'Slide')]"))
                    {
                        var className = node.Attributes["class"].Value;
                        var slideNumber = int.Parse(className.Replace("Slide", ""));
                        maxSlideNumber = Math.Max(maxSlideNumber, slideNumber);
                    }

                    while (_presentation.Slides.Count < maxSlideNumber)
                    {
                        _presentation.Slides.Add(_presentation.Slides.Count + 1, PpSlideLayout.ppLayoutBlank);
                        var newSlide = _presentation.Slides[_presentation.Slides.Count];
                        var newSlideDiv = doc.DocumentNode.SelectSingleNode($"//div[contains(@class, 'Slide{newSlide.SlideNumber}')]");
                        
                        if (newSlideDiv != null)
                        {
                            _slide = newSlide;
                            foreach (var node in newSlideDiv.ChildNodes)
                            {
                                if (node.NodeType == HtmlNodeType.Element)
                                {
                                    ProcessHtmlNode(node);
                                }
                            }
                        }
                    }
                    return true;
                }

                // 단일 슬라이드 처리
                if (_slide == null)
                {
                    Console.WriteLine("현재 슬라이드가 없습니다.");
                    return false;
                }

                // 슬라이드 번호 처리
                if (lineNumber.StartsWith("Slide "))
                {
                    int slideNumber = int.Parse(lineNumber.Replace("Slide ", ""));
                    if (slideNumber > 0)
                    {
                        // 필요한 만큼 슬라이드 생성
                        while (_presentation.Slides.Count < slideNumber)
                        {
                            _presentation.Slides.Add(_presentation.Slides.Count + 1, PpSlideLayout.ppLayoutBlank);
                        }
                        _slide = _presentation.Slides[slideNumber];
                    }
                    else
                    {
                        Console.WriteLine($"슬라이드 번호가 유효하지 않습니다: {slideNumber}");
                        return false;
                    }
                }

                // 기존 도형 삭제
                while (_slide.Shapes.Count > 0)
                {
                    _slide.Shapes[1].Delete();
                }

                var singleDoc = new HtmlDocument();
                singleDoc.LoadHtml(text);

                // HTML 요소를 PowerPoint 도형으로 변환
                foreach (var node in singleDoc.DocumentNode.ChildNodes)
                {
                    if (node.NodeType == HtmlNodeType.Element)
                    {
                        ProcessHtmlNode(node);
                    }
                }

                Console.WriteLine("HTML 텍스트 적용 완료");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"HTML 텍스트 적용 오류: {ex.Message}");
                return false;
            }
        }

        private void ProcessHtmlNode(HtmlNode node)
        {
            if (node.Name.ToLower() == "img")
            {
                try
                {
                    // 이미지 스타일 파싱
                    var style = node.Attributes["style"]?.Value ?? "";
                    var styleDict = new Dictionary<string, string>();
                    foreach (var stylePart in style.Split(';'))
                    {
                        var parts = stylePart.Split(':');
                        if (parts.Length == 2)
                        {
                            styleDict[parts[0].Trim()] = parts[1].Trim();
                        }
                    }

                    // 이미지 소스 가져오기
                    var src = node.Attributes["src"]?.Value;
                    if (string.IsNullOrEmpty(src))
                    {
                        Console.WriteLine("이미지 소스가 없습니다.");
                        return;
                    }

                    // 이미지 파일 경로 생성
                    string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, src);
                    if (!File.Exists(imagePath))
                    {
                        Console.WriteLine($"이미지 파일을 찾을 수 없습니다: {imagePath}");
                        return;
                    }

                    // 이미지 추가
                    var shape = _slide.Shapes.AddPicture(
                        imagePath,
                        MsoTriState.msoFalse,
                        MsoTriState.msoTrue,
                        0, 0);

                    // 스타일 적용 (안전한 파싱)
                    try
                    {
                        if (styleDict.ContainsKey("left"))
                        {
                            var leftValue = ParseSafeFloat(styleDict["left"]);
                            if (IsValidPosition(leftValue))
                                shape.Left = leftValue;
                        }
                        if (styleDict.ContainsKey("top"))
                        {
                            var topValue = ParseSafeFloat(styleDict["top"]);
                            if (IsValidPosition(topValue))
                                shape.Top = topValue;
                        }
                        if (styleDict.ContainsKey("width"))
                        {
                            var widthValue = ParseSafeFloat(styleDict["width"], 100);
                            if (IsValidSize(widthValue))
                                shape.Width = widthValue;
                        }
                        if (styleDict.ContainsKey("height"))
                        {
                            var heightValue = ParseSafeFloat(styleDict["height"], 100);
                            if (IsValidSize(heightValue))
                                shape.Height = heightValue;
                        }
                        if (styleDict.ContainsKey("z-index"))
                        {
                            var zIndex = (int)ParseSafeFloat(styleDict["z-index"]);
                            if (zIndex >= 0 && zIndex <= 100)
                            {
                                shape.ZOrder(MsoZOrderCmd.msoBringToFront);
                                for (int i = 0; i < zIndex; i++)
                                {
                                    shape.ZOrder(MsoZOrderCmd.msoSendBackward);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"이미지 스타일 적용 중 오류: {ex.Message}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"이미지 처리 오류: {ex.Message}");
                }
                return;
            }

            if (node.Name.ToLower() == "div")
            {
                var shape = _slide.Shapes.AddShape(
                    MsoAutoShapeType.msoShapeRectangle,
                    0, 0, 100, 50);

                // div의 스타일 적용
                if (node.Attributes["style"] != null)
                {
                    var divStyle = node.Attributes["style"].Value;
                    var styleDict = new Dictionary<string, string>();
                    foreach (var stylePart in divStyle.Split(';'))
                    {
                        var parts = stylePart.Split(':');
                        if (parts.Length == 2)
                        {
                            styleDict[parts[0].Trim()] = parts[1].Trim();
                        }
                    }

                    // 위치와 크기 적용 (안전한 파싱)
                    try
                    {
                        if (styleDict.ContainsKey("left"))
                        {
                            var leftValue = ParseSafeFloat(styleDict["left"]);
                            if (IsValidPosition(leftValue))
                                shape.Left = leftValue;
                        }
                        if (styleDict.ContainsKey("top"))
                        {
                            var topValue = ParseSafeFloat(styleDict["top"]);
                            if (IsValidPosition(topValue))
                                shape.Top = topValue;
                        }
                        if (styleDict.ContainsKey("width"))
                        {
                            var widthValue = ParseSafeFloat(styleDict["width"], 100);
                            if (IsValidSize(widthValue))
                                shape.Width = widthValue;
                        }
                        if (styleDict.ContainsKey("height"))
                        {
                            var heightValue = ParseSafeFloat(styleDict["height"], 50);
                            if (IsValidSize(heightValue))
                                shape.Height = heightValue;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"div 위치/크기 적용 중 오류: {ex.Message}");
                    }

                    // 텍스트 정렬 적용
                    if (styleDict.ContainsKey("text-align"))
                    {
                        var textAlign = styleDict["text-align"];
                        switch (textAlign)
                        {
                            case "center":
                                shape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                                break;
                            case "right":
                                shape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignRight;
                                break;
                            case "justify":
                                shape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignJustify;
                                break;
                            default:
                                shape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
                                break;
                        }
                    }

                    // 수직 정렬 적용
                    if (styleDict.ContainsKey("vertical-align"))
                    {
                        var verticalAlign = styleDict["vertical-align"];
                        switch (verticalAlign)
                        {
                            case "middle":
                                shape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                                break;
                            case "bottom":
                                shape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorBottom;
                                break;
                            default:
                                shape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorTop;
                                break;
                        }
                    }

                    // 기존 스타일 적용 메서드 호출
                    ApplyStyleToShape(shape, divStyle);
                }

                // span 태그 처리
                foreach (var spanNode in node.ChildNodes)
                {
                    if (spanNode.NodeType == HtmlNodeType.Element && spanNode.Name.ToLower() == "span")
                    {
                        Console.WriteLine($"Span 태그 발견: {spanNode.OuterHtml}");
                        
                        // span의 스타일 적용
                        if (spanNode.Attributes["style"] != null)
                        {
                            var spanStyle = spanNode.Attributes["style"].Value;
                            Console.WriteLine($"Span 스타일: {spanStyle}");
                            var styleDict = new Dictionary<string, string>();
                            foreach (var stylePart in spanStyle.Split(';'))
                            {
                                var parts = stylePart.Split(':');
                                if (parts.Length == 2)
                                {
                                    styleDict[parts[0].Trim()] = parts[1].Trim();
                                }
                            }

                            // 텍스트 설정
                            if (spanNode.InnerText != null)
                            {
                                Console.WriteLine($"Span 텍스트: {spanNode.InnerText}");
                                shape.TextFrame.TextRange.Text = spanNode.InnerText;
                                
                                // <s> 태그 확인
                                if (spanNode.InnerHtml.Contains("<s>"))
                                {
                                    Console.WriteLine("취소선 적용");
                                    if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame2.HasText == MsoTriState.msoTrue)
                                    {
                                        shape.TextFrame2.TextRange.Font.StrikeThrough = MsoTriState.msoTrue;
                                    }
                                }
                            }

                            // 텍스트 스타일 적용
                            if (styleDict.ContainsKey("font-weight"))
                            {
                                Console.WriteLine($"굵게 적용: {styleDict["font-weight"]}");
                                shape.TextFrame.TextRange.Font.Bold = styleDict["font-weight"] == "bold" ? MsoTriState.msoTrue : MsoTriState.msoFalse;
                            }
                            if (styleDict.ContainsKey("font-style"))
                            {
                                Console.WriteLine($"기울임 적용: {styleDict["font-style"]}");
                                shape.TextFrame.TextRange.Font.Italic = styleDict["font-style"] == "italic" ? MsoTriState.msoTrue : MsoTriState.msoFalse;
                            }
                            if (styleDict.ContainsKey("text-decoration"))
                            {
                                var decoration = styleDict["text-decoration"];
                                Console.WriteLine($"텍스트 장식 적용: {decoration}");
                                if (decoration.Contains("underline"))
                                {
                                    Console.WriteLine("밑줄 적용");
                                    if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame2.HasText == MsoTriState.msoTrue)
                                    {
                                        shape.TextFrame2.TextRange.Font.UnderlineStyle = MsoTextUnderlineType.msoUnderlineSingleLine;
                                    }
                                    else
                                    {
                                        shape.TextFrame.TextRange.Font.Underline = MsoTriState.msoTrue;
                                    }
                                }
                                if (decoration.Contains("line-through"))
                                {
                                    Console.WriteLine("취소선 적용");
                                    if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame2.HasText == MsoTriState.msoTrue)
                                    {
                                        shape.TextFrame2.TextRange.Font.StrikeThrough = MsoTriState.msoTrue;
                                    }
                                }
                            }

                            // 폰트 크기 적용 (안전한 파싱)
                            try
                            {
                                if (styleDict.ContainsKey("font-size"))
                                {
                                    var fontSize = ParseSafeFloat(styleDict["font-size"], 11);
                                    if (fontSize >= 1 && fontSize <= 1638)
                                    {
                                        shape.TextFrame.TextRange.Font.Size = fontSize;
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"span 폰트 크기 적용 중 오류: {ex.Message}");
                            }

                            // 하이라이트 색상 적용
                            if (styleDict.ContainsKey("background-color"))
                            {
                                var bgColor = styleDict["background-color"];
                                if (bgColor.StartsWith("rgba"))
                                {
                                    var rgba = bgColor.Replace("rgba(", "").Replace(")", "").Split(',');
                                    if (rgba.Length == 4)
                                    {
                                        int r = int.Parse(rgba[0].Trim());
                                        int g = int.Parse(rgba[1].Trim());
                                        int b = int.Parse(rgba[2].Trim());
                                        float a = float.Parse(rgba[3].Trim());
                                        
                                        if (a > 0.1f)
                                        {
                                            shape.TextFrame2.TextRange.Font.Highlight.RGB = (b << 16) | (g << 8) | r;
                                        }
                                    }
                                }
                                else if (bgColor.StartsWith("rgb"))
                                {
                                    var rgb = bgColor.Replace("rgb(", "").Replace(")", "").Split(',');
                                    if (rgb.Length == 3)
                                    {
                                        int r = int.Parse(rgb[0].Trim());
                                        int g = int.Parse(rgb[1].Trim());
                                        int b = int.Parse(rgb[2].Trim());
                                        shape.TextFrame2.TextRange.Font.Highlight.RGB = (b << 16) | (g << 8) | r;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        public (ulong? FileId, uint? VolumeId, string FileType, string FileName, string FilePath) GetFileInfo()
        {
            try
            {
                if (_presentation == null)
                    return (null, null, "PowerPoint", string.Empty, string.Empty);

                string filePath = _presentation.FullName;
                if (string.IsNullOrEmpty(filePath))
                    return (null, null, "PowerPoint", _presentation.Name, string.Empty);

                return (
                    null,
                    null,
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
