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

<<<<<<< Updated upstream
            // 위치와 크기 적용 (안전한 파싱과 검증)
            try
=======
            // 위치와 크기 적용 (안전한 파싱과 유효성 검사)
            if (styleDict.ContainsKey("left"))
            {
                var left = ParseSafeFloat(styleDict["left"].Replace("px", ""), 0);
                if (left >= 0 && left <= 10000) // PowerPoint 유효 범위
                    shape.Left = left;
            }
            if (styleDict.ContainsKey("top"))
            {
                var top = ParseSafeFloat(styleDict["top"].Replace("px", ""), 0);
                if (top >= 0 && top <= 10000) // PowerPoint 유효 범위
                    shape.Top = top;
            }
            if (styleDict.ContainsKey("width"))
            {
                var width = ParseSafeFloat(styleDict["width"].Replace("px", ""), 100);
                if (width >= 1 && width <= 5000) // PowerPoint 유효 범위
                    shape.Width = width;
            }
            if (styleDict.ContainsKey("height"))
            {
                var height = ParseSafeFloat(styleDict["height"].Replace("px", ""), 50);
                if (height >= 1 && height <= 5000) // PowerPoint 유효 범위
                    shape.Height = height;
            }

            // 회전 적용
            if (styleDict.ContainsKey("transform"))
>>>>>>> Stashed changes
            {
                if (styleDict.ContainsKey("left"))
                {
<<<<<<< Updated upstream
                    var leftValue = ParseSafeFloat(styleDict["left"]);
                    if (IsValidPosition(leftValue))
                        shape.Left = leftValue;
=======
                    var rotation = transform.Replace("rotate(", "").Replace("deg)", "");
                    var rotationValue = ParseSafeFloat(rotation, 0);
                    if (rotationValue >= -360 && rotationValue <= 360) // 회전 유효 범위
                        shape.Rotation = rotationValue;
>>>>>>> Stashed changes
                }
                if (styleDict.ContainsKey("top"))
                {
<<<<<<< Updated upstream
                    var topValue = ParseSafeFloat(styleDict["top"]);
                    if (IsValidPosition(topValue))
                        shape.Top = topValue;
=======
                    var rotationX = transform.Replace("rotateX(", "").Replace("deg)", "");
                    var rotationXValue = ParseSafeFloat(rotationX, 0);
                    if (rotationXValue >= -360 && rotationXValue <= 360) // 회전 유효 범위
                        shape.ThreeD.RotationX = rotationXValue;
>>>>>>> Stashed changes
                }
                if (styleDict.ContainsKey("width"))
                {
<<<<<<< Updated upstream
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
=======
                    var rotationY = transform.Replace("rotateY(", "").Replace("deg)", "");
                    var rotationYValue = ParseSafeFloat(rotationY, 0);
                    if (rotationYValue >= -360 && rotationYValue <= 360) // 회전 유효 범위
                        shape.ThreeD.RotationY = rotationYValue;
                }
            }

            // 3D 효과 적용 (안전한 처리)
            try
            {
                if (styleDict.ContainsKey("transform-style") && styleDict["transform-style"] == "preserve-3d")
                {
                    shape.ThreeD.Visible = MsoTriState.msoTrue;
                    Console.WriteLine("3D 효과 활성화");
                }
                if (styleDict.ContainsKey("perspective"))
                {
                    var perspective = styleDict["perspective"].Replace("px", "");
                    var perspectiveValue = ParseSafeFloat(perspective, 0);
                    if (perspectiveValue >= 0 && perspectiveValue <= 1000) // 원근감 유효 범위
                    {
                        shape.ThreeD.Perspective = (MsoTriState)perspectiveValue;
                        Console.WriteLine($"원근감 설정: {perspectiveValue}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"3D 효과 적용 오류 (무시하고 계속): {ex.Message}");
>>>>>>> Stashed changes
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
                        float a = ParseSafeFloat(rgba[3].Trim(), 1.0f);
                        
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
<<<<<<< Updated upstream
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
=======
                    var weight = ParseSafeFloat(border[0].Replace("px", ""), 1.0f);
                    if (weight >= 0 && weight <= 100) // 테두리 두께 유효 범위
                        shape.Line.Weight = weight;
                    shape.Line.ForeColor.RGB = ParseColor(border[2]);
>>>>>>> Stashed changes
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"테두리 적용 중 오류: {ex.Message}");
                shape.Line.Visible = MsoTriState.msoFalse;
            }

<<<<<<< Updated upstream
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
=======
            // 모서리 둥글기 적용 (일부 도형만 지원)
            if (styleDict.ContainsKey("border-radius"))
            {
                try
                {
                    var radius = styleDict["border-radius"].Replace("px", "");
                    var radiusValue = ParseSafeFloat(radius, 0);
                    if (radiusValue >= 0 && radiusValue <= 100) // 모서리 둥글기 유효 범위
                    {
                        // Adjustments 속성이 있는 도형만 처리
                        if (shape.Adjustments.Count > 0)
                        {
                            shape.Adjustments[1] = radiusValue;
                            Console.WriteLine($"모서리 둥글기 적용: {radiusValue}");
                        }
                        else
                        {
                            Console.WriteLine("이 도형은 Adjustments를 지원하지 않아 모서리 둥글기를 건너뜁니다.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"모서리 둥글기 적용 오류 (무시하고 계속): {ex.Message}");
                }
            }

            // 그림자 효과 적용
            if (styleDict.ContainsKey("box-shadow"))
            {
                var shadow = styleDict["box-shadow"].Split(' ');
                if (shadow.Length >= 4)
                {
                    shape.Shadow.Visible = MsoTriState.msoTrue;
                    var offsetX = ParseSafeFloat(shadow[0].Replace("px", ""), 0);
                    var offsetY = ParseSafeFloat(shadow[1].Replace("px", ""), 0);
                    var blur = ParseSafeFloat(shadow[2].Replace("px", ""), 0);
                    
                    if (offsetX >= -100 && offsetX <= 100) // 그림자 오프셋 유효 범위
                        shape.Shadow.OffsetX = offsetX;
                    if (offsetY >= -100 && offsetY <= 100) // 그림자 오프셋 유효 범위
                        shape.Shadow.OffsetY = offsetY;
                    if (blur >= 0 && blur <= 100) // 그림자 흐림 유효 범위
                        shape.Shadow.Blur = blur;
                    shape.Shadow.ForeColor.RGB = ParseColor(shadow[3]);
                }
            }

            // Z-인덱스 적용 (안전한 처리)
            if (styleDict.ContainsKey("z-index"))
            {
                try
                {
                    if (int.TryParse(styleDict["z-index"], out int zIndex) && zIndex >= 0 && zIndex <= 100)
                    {
                        // Z-인덱스는 읽기 전용이므로 ZOrder 메서드를 사용
                        shape.ZOrder(MsoZOrderCmd.msoBringToFront);
                        for (int i = 0; i < zIndex; i++)
                        {
                            shape.ZOrder(MsoZOrderCmd.msoSendBackward);
                        }
                        Console.WriteLine($"Z-인덱스 적용: {zIndex}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Z-인덱스 적용 오류 (무시하고 계속): {ex.Message}");
>>>>>>> Stashed changes
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

        private float ParseSafeFloat(string value, float defaultValue)
        {
            if (string.IsNullOrWhiteSpace(value))
                return defaultValue;
                
            if (float.TryParse(value, out float result))
                return result;
                
            return defaultValue;
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
                Console.WriteLine($"라인 번호 파라미터: '{lineNumber}'");
                
                // 슬라이드 번호 처리
                if (lineNumber.StartsWith("Slide "))
                {
                    string slideNumberStr = lineNumber.Replace("Slide ", "").Trim();
                    Console.WriteLine($"슬라이드 번호 문자열: '{slideNumberStr}'");
                    
                    if (int.TryParse(slideNumberStr, out int slideNumber) && slideNumber > 0)
                    {
                        Console.WriteLine($"적용할 슬라이드 번호: {slideNumber}");
                        Console.WriteLine($"현재 프레젠테이션의 슬라이드 수: {_presentation.Slides.Count}");
                        
                        // 필요한 만큼 슬라이드 생성
                        while (_presentation.Slides.Count < slideNumber)
                        {
                            _presentation.Slides.Add(_presentation.Slides.Count + 1, PpSlideLayout.ppLayoutBlank);
                            Console.WriteLine($"새 슬라이드 생성됨. 현재 슬라이드 수: {_presentation.Slides.Count}");
                        }
                        
                        _slide = _presentation.Slides[slideNumber];
                        Console.WriteLine($"슬라이드 {slideNumber} 선택 완료");
                    }
                    else
                    {
                        Console.WriteLine($"슬라이드 번호 파싱 실패: '{slideNumberStr}'");
                        return false;
                    }
                }
                else
                {
                    // lineNumber가 "Slide " 형태가 아닌 경우 첫 번째 슬라이드 사용
                    Console.WriteLine($"슬라이드 번호 형태가 아닙니다. 첫 번째 슬라이드를 사용합니다: '{lineNumber}'");
                    if (_presentation.Slides.Count > 0)
                    {
                        _slide = _presentation.Slides[1];
                        Console.WriteLine("첫 번째 슬라이드 선택 완료");
                    }
                    else
                    {
                        Console.WriteLine("슬라이드가 없습니다. 새 슬라이드를 생성합니다.");
                        _presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank);
                        _slide = _presentation.Slides[1];
                        Console.WriteLine("새 슬라이드 생성 및 선택 완료");
                    }
                }

                if (_slide == null)
                {
                    Console.WriteLine("슬라이드 선택 실패");
                    return false;
                }

                // 기존 도형 삭제
                Console.WriteLine($"선택된 슬라이드의 기존 도형 수: {_slide.Shapes.Count}");
                int deletedShapeCount = 0;
                while (_slide.Shapes.Count > 0)
                {
                    _slide.Shapes[1].Delete();
                    deletedShapeCount++;
                }
                Console.WriteLine($"삭제된 도형 수: {deletedShapeCount}");

                var singleDoc = new HtmlDocument();
                singleDoc.LoadHtml(text);
                
                Console.WriteLine($"HTML 콘텐츠 로드 완료. 루트 노드 수: {singleDoc.DocumentNode.ChildNodes.Count}");
                Console.WriteLine($"HTML 콘텐츠 미리보기: {text.Substring(0, Math.Min(200, text.Length))}");

                // HTML 요소를 PowerPoint 도형으로 변환
                int processedNodes = 0;
                foreach (var node in singleDoc.DocumentNode.ChildNodes)
                {
                    if (node.NodeType == HtmlNodeType.Element)
                    {
                        Console.WriteLine($"HTML 노드 처리 중: {node.Name}");
                        ProcessHtmlNode(node);
                        processedNodes++;
                    }
                }
                Console.WriteLine($"처리된 HTML 노드 수: {processedNodes}");

                Console.WriteLine($"HTML 텍스트 적용 완료. 최종 슬라이드 도형 수: {_slide.Shapes.Count}");
                Console.WriteLine($"적용된 슬라이드 번호: {_slide.SlideNumber}");
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
            try
            {
                Console.WriteLine($"ProcessHtmlNode 시작: {node.Name}");
                Console.WriteLine($"현재 슬라이드 정보: SlideNumber={_slide?.SlideNumber}, 도형 수={_slide?.Shapes.Count}");
                
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

<<<<<<< Updated upstream
                    // 스타일 적용 (안전한 파싱)
                    try
                    {
                        if (styleDict.ContainsKey("left"))
                        {
                            var leftValue = ParseSafeFloat(styleDict["left"]);
                            if (IsValidPosition(leftValue))
                                shape.Left = leftValue;
=======
                    // 스타일 적용 (안전한 파싱과 유효성 검사)
                    if (styleDict.ContainsKey("left"))
                    {
                        var left = ParseSafeFloat(styleDict["left"].Replace("px", ""), 0);
                        if (left >= 0 && left <= 10000) // PowerPoint 유효 범위
                            shape.Left = left;
                    }
                    if (styleDict.ContainsKey("top"))
                    {
                        var top = ParseSafeFloat(styleDict["top"].Replace("px", ""), 0);
                        if (top >= 0 && top <= 10000) // PowerPoint 유효 범위
                            shape.Top = top;
                    }
                    if (styleDict.ContainsKey("width"))
                    {
                        var width = ParseSafeFloat(styleDict["width"].Replace("px", ""), 100);
                        if (width >= 1 && width <= 5000) // PowerPoint 유효 범위
                            shape.Width = width;
                    }
                    if (styleDict.ContainsKey("height"))
                    {
                        var height = ParseSafeFloat(styleDict["height"].Replace("px", ""), 50);
                        if (height >= 1 && height <= 5000) // PowerPoint 유효 범위
                            shape.Height = height;
                    }
                    if (styleDict.ContainsKey("z-index"))
                    {
                        if (int.TryParse(styleDict["z-index"], out int zIndex) && zIndex >= 0 && zIndex <= 100)
                        {
                            shape.ZOrder(MsoZOrderCmd.msoBringToFront);
                            for (int i = 0; i < zIndex; i++)
                            {
                                shape.ZOrder(MsoZOrderCmd.msoSendBackward);
                            }
>>>>>>> Stashed changes
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
                Console.WriteLine("DIV 태그 처리 시작");
                Console.WriteLine($"슬라이드 처리 전 도형 수: {_slide.Shapes.Count}");
                
                var shape = _slide.Shapes.AddShape(
                    MsoAutoShapeType.msoShapeRectangle,
                    0, 0, 100, 50);
                    
                Console.WriteLine($"DIV 도형 생성 완료. 슬라이드 도형 수: {_slide.Shapes.Count}");

                // div의 스타일 적용
                Console.WriteLine($"DIV 태그 내용: {node.InnerText?.Substring(0, Math.Min(50, node.InnerText?.Length ?? 0))}");
                Console.WriteLine($"DIV HTML: {node.OuterHtml.Substring(0, Math.Min(100, node.OuterHtml.Length))}");
                
                if (node.Attributes["style"] != null)
                {
                    var divStyle = node.Attributes["style"].Value;
                    Console.WriteLine($"DIV 스타일: {divStyle}");
                    var styleDict = new Dictionary<string, string>();
                    foreach (var stylePart in divStyle.Split(';'))
                    {
                        var parts = stylePart.Split(':');
                        if (parts.Length == 2)
                        {
                            styleDict[parts[0].Trim()] = parts[1].Trim();
                        }
                    }

<<<<<<< Updated upstream
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
=======
                    // 위치와 크기 적용 (안전한 파싱과 유효성 검사)
                    if (styleDict.ContainsKey("left"))
                    {
                        var left = ParseSafeFloat(styleDict["left"].Replace("px", ""), 0);
                        if (left >= 0 && left <= 10000) // PowerPoint 유효 범위
                            shape.Left = left;
                    }
                    if (styleDict.ContainsKey("top"))
                    {
                        var top = ParseSafeFloat(styleDict["top"].Replace("px", ""), 0);
                        if (top >= 0 && top <= 10000) // PowerPoint 유효 범위
                            shape.Top = top;
                    }
                    if (styleDict.ContainsKey("width"))
                    {
                        var width = ParseSafeFloat(styleDict["width"].Replace("px", ""), 100);
                        if (width >= 1 && width <= 5000) // PowerPoint 유효 범위
                            shape.Width = width;
                    }
                    if (styleDict.ContainsKey("height"))
                    {
                        var height = ParseSafeFloat(styleDict["height"].Replace("px", ""), 50);
                        if (height >= 1 && height <= 5000) // PowerPoint 유효 범위
                            shape.Height = height;
>>>>>>> Stashed changes
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
                    try
                    {
                        ApplyStyleToShape(shape, divStyle);
                        Console.WriteLine("DIV 스타일 적용 완료");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"DIV 스타일 적용 오류 (계속 진행): {ex.Message}");
                    }
                }

                // span 태그 처리
                Console.WriteLine($"DIV의 자식 노드 수: {node.ChildNodes.Count}");
                foreach (var spanNode in node.ChildNodes)
                {
                    Console.WriteLine($"자식 노드 처리: {spanNode.NodeType} - {spanNode.Name}");
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
                                Console.WriteLine($"Span 텍스트: '{spanNode.InnerText}'");
                                shape.TextFrame.TextRange.Text = spanNode.InnerText;
                                Console.WriteLine($"도형에 텍스트 설정 완료: '{shape.TextFrame.TextRange.Text}'");
                                
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
                                        if (int.TryParse(rgba[0].Trim(), out int r) &&
                                            int.TryParse(rgba[1].Trim(), out int g) &&
                                            int.TryParse(rgba[2].Trim(), out int b) &&
                                            float.TryParse(rgba[3].Trim(), out float a))
                                        
                                        {
                                            if (a > 0.1f)
                                            {
                                                shape.TextFrame2.TextRange.Font.Highlight.RGB = (b << 16) | (g << 8) | r;
                                            }
                                        }
                                    }
                                }
                                else if (bgColor.StartsWith("rgb"))
                                {
                                    var rgb = bgColor.Replace("rgb(", "").Replace(")", "").Split(',');
                                    if (rgb.Length == 3)
                                    {
                                        if (int.TryParse(rgb[0].Trim(), out int r) &&
                                            int.TryParse(rgb[1].Trim(), out int g) &&
                                            int.TryParse(rgb[2].Trim(), out int b))
                                        {
                                            shape.TextFrame2.TextRange.Font.Highlight.RGB = (b << 16) | (g << 8) | r;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                
                // span이 없고 div에 직접 텍스트가 있는 경우 처리
                if (!string.IsNullOrWhiteSpace(node.InnerText))
                {
                    bool hasSpanChild = false;
                    foreach (var childNode in node.ChildNodes)
                    {
                        if (childNode.NodeType == HtmlNodeType.Element && childNode.Name.ToLower() == "span")
                        {
                            hasSpanChild = true;
                            break;
                        }
                    }
                    
                    if (!hasSpanChild)
                    {
                        Console.WriteLine($"DIV에 직접 텍스트 설정: '{node.InnerText}'");
                        shape.TextFrame.TextRange.Text = node.InnerText;
                        Console.WriteLine($"DIV 도형에 텍스트 설정 완료: '{shape.TextFrame.TextRange.Text}'");
                    }
                }
                
                Console.WriteLine($"DIV 태그 처리 완료. 최종 도형 수: {_slide.Shapes.Count}");
            }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ProcessHtmlNode 오류: {ex.Message}");
                var outerHtml = node?.OuterHtml ?? "";
                var maxLength = Math.Min(100, outerHtml.Length);
                Console.WriteLine($"노드 정보: {node?.Name} - {outerHtml.Substring(0, maxLength)}");
                Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                // 오류가 발생해도 계속 진행
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
