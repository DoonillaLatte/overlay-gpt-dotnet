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
using System.Linq;
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
            try
            {
                Console.WriteLine($"스타일 적용 시작: {style.Substring(0, Math.Min(100, style.Length))}");
                
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
                        {
                            shape.Left = leftValue;
                            Console.WriteLine($"Left 적용: {leftValue}");
                        }
                    }
                    if (styleDict.ContainsKey("top"))
                    {
                        var topValue = ParseSafeFloat(styleDict["top"]);
                        if (IsValidPosition(topValue))
                        {
                            shape.Top = topValue;
                            Console.WriteLine($"Top 적용: {topValue}");
                        }
                    }
                    if (styleDict.ContainsKey("width"))
                    {
                        var widthValue = ParseSafeFloat(styleDict["width"], 100);
                        if (IsValidSize(widthValue))
                        {
                            shape.Width = widthValue;
                            Console.WriteLine($"Width 적용: {widthValue}");
                        }
                    }
                    if (styleDict.ContainsKey("height"))
                    {
                        var heightValue = ParseSafeFloat(styleDict["height"], 50);
                        if (IsValidSize(heightValue))
                        {
                            shape.Height = heightValue;
                            Console.WriteLine($"Height 적용: {heightValue}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"위치/크기 적용 중 오류 (무시하고 계속): {ex.Message}");
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
                            {
                                shape.Rotation = rotationValue;
                                Console.WriteLine($"회전 적용: {rotationValue}도");
                            }
                        }
                        else if (transform.Contains("rotateX"))
                        {
                            var rotationX = transform.Replace("rotateX(", "").Replace("deg)", "");
                            var rotationXValue = ParseSafeFloat(rotationX);
                            if (rotationXValue >= -360 && rotationXValue <= 360)
                            {
                                shape.ThreeD.RotationX = rotationXValue;
                                Console.WriteLine($"X축 회전 적용: {rotationXValue}도");
                            }
                        }
                        else if (transform.Contains("rotateY"))
                        {
                            var rotationY = transform.Replace("rotateY(", "").Replace("deg)", "");
                            var rotationYValue = ParseSafeFloat(rotationY);
                            if (rotationYValue >= -360 && rotationYValue <= 360)
                            {
                                shape.ThreeD.RotationY = rotationYValue;
                                Console.WriteLine($"Y축 회전 적용: {rotationYValue}도");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"회전 적용 중 오류 (무시하고 계속): {ex.Message}");
                }

                // 3D 효과 적용 (안전한 파싱)
                try
                {
                    if (styleDict.ContainsKey("transform-style") && styleDict["transform-style"] == "preserve-3d")
                    {
                        shape.ThreeD.Visible = MsoTriState.msoTrue;
                        Console.WriteLine("3D 효과 활성화");
                    }
                    if (styleDict.ContainsKey("perspective"))
                    {
                        var perspectiveValue = ParseSafeFloat(styleDict["perspective"]);
                        if (perspectiveValue >= 0 && perspectiveValue <= 5000)
                        {
                            shape.ThreeD.Perspective = (MsoTriState)perspectiveValue;
                            Console.WriteLine($"원근감 설정: {perspectiveValue}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"3D 효과 적용 오류 (무시하고 계속): {ex.Message}");
                }

                // 텍스트 스타일 적용 (안전한 처리)
                try
                {
                    if (styleDict.ContainsKey("font-family"))
                    {
                        shape.TextFrame.TextRange.Font.Name = styleDict["font-family"];
                        Console.WriteLine($"폰트 적용: {styleDict["font-family"]}");
                    }
                    if (styleDict.ContainsKey("font-weight"))
                    {
                        shape.TextFrame.TextRange.Font.Bold = styleDict["font-weight"] == "bold" ? MsoTriState.msoTrue : MsoTriState.msoFalse;
                        Console.WriteLine($"굵게 적용: {styleDict["font-weight"]}");
                    }
                    if (styleDict.ContainsKey("font-style"))
                    {
                        shape.TextFrame.TextRange.Font.Italic = styleDict["font-style"] == "italic" ? MsoTriState.msoTrue : MsoTriState.msoFalse;
                        Console.WriteLine($"기울임 적용: {styleDict["font-style"]}");
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
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"텍스트 스타일 적용 오류 (무시하고 계속): {ex.Message}");
                }

                // 텍스트 색상 적용 (안전한 처리)
                try
                {
                    if (styleDict.ContainsKey("color"))
                    {
                        var color = styleDict["color"];
                        if (color.StartsWith("#"))
                        {
                            shape.TextFrame.TextRange.Font.Color.RGB = ParseColor(color);
                            Console.WriteLine($"텍스트 색상 적용: {color}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"텍스트 색상 적용 오류 (무시하고 계속): {ex.Message}");
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
                            Console.WriteLine($"폰트 크기 적용: {fontSize}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"폰트 크기 적용 중 오류 (무시하고 계속): {ex.Message}");
                }

                // 배경색 적용 (안전한 처리)
                try
                {
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
                                    Console.WriteLine("배경색 투명화");
                                }
                                else
                                {
                                    shape.Fill.Visible = MsoTriState.msoTrue;
                                    shape.Fill.ForeColor.RGB = (b << 16) | (g << 8) | r;
                                    shape.Fill.Transparency = 1 - a;
                                    Console.WriteLine($"RGBA 배경색 적용: {bgColor}");
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
                                Console.WriteLine($"RGB 배경색 적용: {bgColor}");
                            }
                        }
                        else if (bgColor.StartsWith("#"))
                        {
                            shape.Fill.Visible = MsoTriState.msoTrue;
                            shape.Fill.ForeColor.RGB = ParseColor(bgColor);
                            Console.WriteLine($"HEX 배경색 적용: {bgColor}");
                        }
                    }
                    else
                    {
                        // background-color가 명시되지 않은 경우 투명 배경으로 설정
                        shape.Fill.Visible = MsoTriState.msoFalse;
                        Console.WriteLine("배경색 미명시 - 투명 배경 적용");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"배경색 적용 오류 (무시하고 계속): {ex.Message}");
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
                                Console.WriteLine($"테두리 적용: {borderWeight}px {border[2]}");
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
                    Console.WriteLine($"테두리 적용 중 오류 (무시하고 계속): {ex.Message}");
                    try
                    {
                        shape.Line.Visible = MsoTriState.msoFalse;
                    }
                    catch { }
                }

                // 모서리 둥글기 적용 (안전한 처리 - 일부 도형만 지원)
                try
                {
                    if (styleDict.ContainsKey("border-radius"))
                    {
                        var radius = ParseSafeFloat(styleDict["border-radius"]);
                        if (radius >= 0 && radius <= 100)
                        {
                            // Adjustments 속성이 있는 도형만 처리
                            if (shape.Adjustments.Count > 0)
                            {
                                shape.Adjustments[1] = radius;
                                Console.WriteLine($"모서리 둥글기 적용: {radius}");
                            }
                            else
                            {
                                Console.WriteLine("이 도형은 Adjustments를 지원하지 않아 모서리 둥글기를 건너뜁니다.");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"모서리 둥글기 적용 오류 (무시하고 계속): {ex.Message}");
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
                                Console.WriteLine($"그림자 효과 적용: {offsetX}, {offsetY}, {blur}, {shadow[3]}");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"그림자 효과 적용 중 오류 (무시하고 계속): {ex.Message}");
                }

                // Z-인덱스 적용 (안전한 파싱)
                try
                {
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
                            Console.WriteLine($"Z-인덱스 적용: {zIndex}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Z-인덱스 적용 중 오류 (무시하고 계속): {ex.Message}");
                }

                // 텍스트 정렬 적용 (안전한 처리)
                try
                {
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
                        Console.WriteLine($"텍스트 정렬 적용: {textAlign}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"텍스트 정렬 적용 오류 (무시하고 계속): {ex.Message}");
                }

                // 수직 정렬 적용 (안전한 처리)
                try
                {
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
                        Console.WriteLine($"수직 정렬 적용: {verticalAlign}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"수직 정렬 적용 오류 (무시하고 계속): {ex.Message}");
                }
                
                Console.WriteLine("스타일 적용 완료");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"전체 스타일 적용 오류 (무시하고 계속): {ex.Message}");
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

        private void ProcessHtmlNode(HtmlNode node)
        {
            try
            {
                if (node.NodeType != HtmlNodeType.Element)
                    return;

                Console.WriteLine($"ProcessHtmlNode 시작: {node.Name}");
                Console.WriteLine($"현재 슬라이드 정보: SlideNumber={_slide?.SlideNumber}, 도형 수={_slide?.Shapes.Count}");

                if (node.Name.ToLower() == "div")
                {
                    Console.WriteLine("DIV 태그 처리 시작");
                    Console.WriteLine($"슬라이드 처리 전 도형 수: {_slide.Shapes.Count}");
                    
                    // DIV를 도형으로 변환
                    var shape = _slide.Shapes.AddShape(
                        MsoAutoShapeType.msoShapeRectangle,
                        0, 0, 100, 50);

                    Console.WriteLine($"DIV 도형 생성 완료. 슬라이드 도형 수: {_slide.Shapes.Count}");

                    // 스타일 적용 (안전한 처리)
                    var divStyle = node.GetAttributeValue("style", "");
                    Console.WriteLine($"DIV 스타일: {divStyle}");
                    
                    if (!string.IsNullOrEmpty(divStyle))
                    {
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

                    // 자식 노드 처리 (span 태그 등)
                    Console.WriteLine($"DIV의 자식 노드 수: {node.ChildNodes.Count}");
                    foreach (var childNode in node.ChildNodes)
                    {
                        Console.WriteLine($"자식 노드 처리: {childNode.NodeType} - {childNode.Name}");
                        if (childNode.NodeType == HtmlNodeType.Element)
                        {
                            if (childNode.Name.ToLower() == "span")
                            {
                                Console.WriteLine($"Span 태그 발견: {childNode.OuterHtml}");
                                
                                // span의 스타일 적용
                                if (childNode.Attributes["style"] != null)
                                {
                                    var spanStyle = childNode.Attributes["style"].Value;
                                    Console.WriteLine($"Span 스타일: {spanStyle}");
                                    
                                    try
                                    {
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
                        if (childNode.InnerText != null)
                        {
                            Console.WriteLine($"Span 텍스트: '{childNode.InnerText}'");
                            
                            // <br> 태그를 불릿 포인트로 변환
                            var processedText = childNode.InnerText;
                            var htmlContent = childNode.InnerHtml;
                            
                            Console.WriteLine($"불릿포인트 변환 시도. HTML: {htmlContent}");
                            Console.WriteLine($"InnerText: {processedText}");
                            
                            // <br> 태그가 있으면 불릿 포인트로 변환
                            if (htmlContent.Contains("<br>"))
                            {
                                Console.WriteLine("HTML에서 <br> 태그 발견, 불릿포인트로 변환 시작");
                                
                                // <br> 태그를 줄바꿈으로 변환
                                string textForSplit = System.Text.RegularExpressions.Regex.Replace(htmlContent, @"<br\s*/?>", "\n", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                                // HTML 태그 제거
                                textForSplit = System.Text.RegularExpressions.Regex.Replace(textForSplit, @"<[^>]+>", "");
                                
                                Console.WriteLine($"HTML 태그 제거 후 텍스트: {textForSplit}");
                                
                                // 줄바꿈으로 분할
                                var sentences = textForSplit.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                                Console.WriteLine($"분할된 문장 수: {sentences.Length}");
                                
                                if (sentences.Length > 1)
                                {
                                    processedText = "• " + string.Join("\n\n• ", sentences.Select(s => s.Trim()).Where(s => !string.IsNullOrEmpty(s)));
                                    Console.WriteLine($"불릿 포인트로 변환 완료: '{processedText}'");
                                }
                                else
                                {
                                    Console.WriteLine("분할된 문장이 1개 이하이므로 불릿포인트 변환하지 않음");
                                }
                            }
                            else if (processedText.Contains("\r"))
                            {
                                Console.WriteLine("텍스트에서 \\r 패턴 발견, 불릿포인트로 변환 시작");
                                
                                var sentences = processedText.Split(new string[] { "\r\r", "\r" }, StringSplitOptions.RemoveEmptyEntries);
                                Console.WriteLine($"분할된 문장 수: {sentences.Length}");
                                
                                if (sentences.Length > 1)
                                {
                                    processedText = "• " + string.Join("\n\n• ", sentences.Select(s => s.Trim()).Where(s => !string.IsNullOrEmpty(s)));
                                    Console.WriteLine($"불릿 포인트로 변환 완료: '{processedText}'");
                                }
                            }
                            
                            shape.TextFrame.TextRange.Text = processedText;
                            Console.WriteLine($"도형에 텍스트 설정 완료: '{shape.TextFrame.TextRange.Text}'");
                                            
                                            // <s> 태그 확인 (취소선)
                                            if (childNode.InnerHtml.Contains("<s>"))
                                            {
                                                Console.WriteLine("취소선 적용");
                                                try
                                                {
                                                    if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame2.HasText == MsoTriState.msoTrue)
                                                    {
                                                        shape.TextFrame2.TextRange.Font.StrikeThrough = MsoTriState.msoTrue;
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    Console.WriteLine($"취소선 적용 오류 (무시): {ex.Message}");
                                                }
                                            }
                                        }

                                        // 하이라이트 색상 적용 (안전한 처리)
                                        try
                                        {
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
                                                                Console.WriteLine($"RGBA 하이라이트 적용: {bgColor}");
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
                                                            Console.WriteLine($"RGB 하이라이트 적용: {bgColor}");
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"하이라이트 색상 적용 오류 (무시): {ex.Message}");
                                        }

                                        // 기본 스타일 적용
                                        try
                                        {
                                            ApplyStyleToShape(shape, spanStyle);
                                            Console.WriteLine("Span 스타일 적용 완료");
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"Span 스타일 적용 오류 (무시): {ex.Message}");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"Span 처리 오류 (무시): {ex.Message}");
                                    }
                                }
                            }
                            else
                            {
                                // 다른 자식 요소는 재귀적으로 처리
                                ProcessHtmlNode(childNode);
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
                            try
                            {
                                shape.TextFrame.TextRange.Text = node.InnerText;
                                Console.WriteLine($"DIV 도형에 텍스트 설정 완료: '{shape.TextFrame.TextRange.Text}'");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"DIV 텍스트 설정 오류 (무시): {ex.Message}");
                            }
                        }
                    }
                    
                    Console.WriteLine($"DIV 태그 처리 완료. 최종 도형 수: {_slide.Shapes.Count}");
                }
                else if (node.Name.ToLower() == "span")
                {
                    Console.WriteLine("독립 SPAN 태그 처리 시작");
                    
                    // SPAN을 텍스트 상자로 변환
                    var shape = _slide.Shapes.AddTextbox(
                        MsoTextOrientation.msoTextOrientationHorizontal,
                        0, 0, 100, 50);

                    // 텍스트 설정
                    if (!string.IsNullOrEmpty(node.InnerText))
                    {
                        try
                        {
                            // <br> 태그를 불릿 포인트로 변환
                            var processedText = node.InnerText;
                            var htmlContent = node.InnerHtml;
                            
                            Console.WriteLine($"독립 SPAN 불릿포인트 변환 시도. HTML: {htmlContent}");
                            
                            // <br> 태그가 있으면 불릿 포인트로 변환
                            if (htmlContent.Contains("<br>"))
                            {
                                Console.WriteLine("독립 SPAN에서 <br> 태그 발견, 불릿포인트로 변환 시작");
                                
                                // <br> 태그를 줄바꿈으로 변환
                                string textForSplit = System.Text.RegularExpressions.Regex.Replace(htmlContent, @"<br\s*/?>", "\n", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                                // HTML 태그 제거
                                textForSplit = System.Text.RegularExpressions.Regex.Replace(textForSplit, @"<[^>]+>", "");
                                
                                Console.WriteLine($"독립 SPAN HTML 태그 제거 후: {textForSplit}");
                                
                                // 줄바꿈으로 분할
                                var sentences = textForSplit.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                                Console.WriteLine($"독립 SPAN 분할된 문장 수: {sentences.Length}");
                                
                                if (sentences.Length > 1)
                                {
                                    processedText = "• " + string.Join("\n\n• ", sentences.Select(s => s.Trim()).Where(s => !string.IsNullOrEmpty(s)));
                                    Console.WriteLine($"독립 SPAN 불릿 포인트로 변환 완료: '{processedText}'");
                                }
                            }
                            else if (processedText.Contains("\r"))
                            {
                                Console.WriteLine("독립 SPAN 텍스트에서 \\r 패턴 발견");
                                
                                var sentences = processedText.Split(new string[] { "\r\r", "\r" }, StringSplitOptions.RemoveEmptyEntries);
                                
                                if (sentences.Length > 1)
                                {
                                    processedText = "• " + string.Join("\n\n• ", sentences.Select(s => s.Trim()).Where(s => !string.IsNullOrEmpty(s)));
                                    Console.WriteLine($"독립 SPAN 불릿 포인트로 변환 완료: '{processedText}'");
                                }
                            }
                            
                            shape.TextFrame.TextRange.Text = processedText;
                            Console.WriteLine($"Span 텍스트 설정: '{processedText}'");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Span 텍스트 설정 오류 (무시): {ex.Message}");
                        }
                    }

                    // 스타일 적용
                    var spanStyle = node.GetAttributeValue("style", "");
                    if (!string.IsNullOrEmpty(spanStyle))
                    {
                        try
                        {
                            ApplyStyleToShape(shape, spanStyle);
                            Console.WriteLine("독립 Span 스타일 적용 완료");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"독립 Span 스타일 적용 오류 (무시): {ex.Message}");
                        }
                    }
                }
                else if (node.Name.ToLower() == "img")
                {
                    Console.WriteLine("IMG 태그 처리 시작");
                    
                    // 이미지 URL 가져오기
                    var imgSrc = node.GetAttributeValue("src", "");
                    if (!string.IsNullOrEmpty(imgSrc))
                    {
                        try
                        {
                            // 이미지 추가 (파일 경로 처리)
                            var shape = _slide.Shapes.AddPicture(
                                imgSrc,
                                Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue,
                                0, 0);

                            Console.WriteLine($"이미지 추가 완료: {imgSrc}");

                            // 스타일 적용
                            var imgStyle = node.GetAttributeValue("style", "");
                            if (!string.IsNullOrEmpty(imgStyle))
                            {
                                try
                                {
                                    ApplyStyleToShape(shape, imgStyle);
                                    Console.WriteLine("이미지 스타일 적용 완료");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"이미지 스타일 적용 오류 (무시): {ex.Message}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"이미지 처리 중 오류: {ex.Message}");
                        }
                    }
                    else
                    {
                        Console.WriteLine("이미지 소스가 없습니다.");
                    }
                }
                else
                {
                    Console.WriteLine($"알 수 없는 태그: {node.Name} - 자식 노드들을 처리합니다.");
                    // 다른 태그의 자식 노드들 처리
                    foreach (var childNode in node.ChildNodes)
                    {
                        if (childNode.NodeType == HtmlNodeType.Element)
                        {
                            ProcessHtmlNode(childNode);
                        }
                    }
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

                // 슬라이드 번호 처리
                if (lineNumber.StartsWith("Slide "))
                {
                    string slideNumberStr = lineNumber.Replace("Slide ", "").Trim();
                    
                    if (int.TryParse(slideNumberStr, out int slideNumber) && slideNumber > 0)
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
                        Console.WriteLine($"슬라이드 번호 파싱 실패: '{slideNumberStr}'");
                        return false;
                    }
                }
                else
                {
                    // 첫 번째 슬라이드 사용
                    if (_presentation.Slides.Count > 0)
                    {
                        _slide = _presentation.Slides[1];
                    }
                    else
                    {
                        _presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank);
                        _slide = _presentation.Slides[1];
                    }
                }

                if (_slide == null)
                {
                    Console.WriteLine("슬라이드 선택 실패");
                    return false;
                }

                // 기존 도형 삭제
                while (_slide.Shapes.Count > 0)
                {
                    _slide.Shapes[1].Delete();
                }

                // HTML 파싱 및 처리
                var doc = new HtmlDocument();
                doc.LoadHtml(text);

                foreach (var node in doc.DocumentNode.ChildNodes)
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

        public (ulong? FileId, uint? VolumeId, string FileType, string FileName, string FilePath) GetFileInfo()
        {
            if (_presentation == null)
                return (null, null, "", "", "");

            try
            {
                var filePath = _presentation.FullName;
                var fileName = Path.GetFileName(filePath);
                return (null, null, "PowerPoint", fileName, filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"파일 정보 가져오기 오류: {ex.Message}");
                return (null, null, "", "", "");
            }
        }
    }
}

