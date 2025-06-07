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

            // 위치와 크기 적용
            if (styleDict.ContainsKey("left"))
                shape.Left = float.Parse(styleDict["left"].Replace("px", ""));
            if (styleDict.ContainsKey("top"))
                shape.Top = float.Parse(styleDict["top"].Replace("px", ""));
            if (styleDict.ContainsKey("width"))
                shape.Width = float.Parse(styleDict["width"].Replace("px", ""));
            if (styleDict.ContainsKey("height"))
                shape.Height = float.Parse(styleDict["height"].Replace("px", ""));

            // 회전 적용
            if (styleDict.ContainsKey("transform"))
            {
                var transform = styleDict["transform"];
                if (transform.Contains("rotate"))
                {
                    var rotation = transform.Replace("rotate(", "").Replace("deg)", "");
                    shape.Rotation = float.Parse(rotation);
                }
                else if (transform.Contains("rotateX"))
                {
                    var rotationX = transform.Replace("rotateX(", "").Replace("deg)", "");
                    shape.ThreeD.RotationX = float.Parse(rotationX);
                }
                else if (transform.Contains("rotateY"))
                {
                    var rotationY = transform.Replace("rotateY(", "").Replace("deg)", "");
                    shape.ThreeD.RotationY = float.Parse(rotationY);
                }
            }

            // 3D 효과 적용
            if (styleDict.ContainsKey("transform-style") && styleDict["transform-style"] == "preserve-3d")
            {
                shape.ThreeD.Visible = MsoTriState.msoTrue;
            }
            if (styleDict.ContainsKey("perspective"))
            {
                var perspective = styleDict["perspective"].Replace("px", "");
                shape.ThreeD.Perspective = (MsoTriState)float.Parse(perspective);
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

            // 폰트 크기 적용
            if (styleDict.ContainsKey("font-size"))
            {
                var fontSize = styleDict["font-size"].Replace("pt", "").Trim();
                if (float.TryParse(fontSize, out float size))
                {
                    shape.TextFrame.TextRange.Font.Size = size;
                }
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

            // 테두리 적용
            if (styleDict.ContainsKey("border"))
            {
                var border = styleDict["border"].Split(' ');
                if (border.Length >= 3)
                {
                    shape.Line.Weight = float.Parse(border[0].Replace("px", ""));
                    shape.Line.ForeColor.RGB = ParseColor(border[2]);
                }
            }
            else
            {
                shape.Line.Visible = MsoTriState.msoFalse;
            }

            // 모서리 둥글기 적용
            if (styleDict.ContainsKey("border-radius"))
            {
                var radius = styleDict["border-radius"].Replace("px", "");
                shape.Adjustments[1] = float.Parse(radius);
            }

            // 그림자 효과 적용
            if (styleDict.ContainsKey("box-shadow"))
            {
                var shadow = styleDict["box-shadow"].Split(' ');
                if (shadow.Length >= 4)
                {
                    shape.Shadow.Visible = MsoTriState.msoTrue;
                    shape.Shadow.OffsetX = float.Parse(shadow[0].Replace("px", ""));
                    shape.Shadow.OffsetY = float.Parse(shadow[1].Replace("px", ""));
                    shape.Shadow.Blur = float.Parse(shadow[2].Replace("px", ""));
                    shape.Shadow.ForeColor.RGB = ParseColor(shadow[3]);
                }
            }

            // Z-인덱스 적용
            if (styleDict.ContainsKey("z-index"))
            {
                var zIndex = int.Parse(styleDict["z-index"]);
                // Z-인덱스는 읽기 전용이므로 ZOrder 메서드를 사용
                shape.ZOrder(MsoZOrderCmd.msoBringToFront);
                for (int i = 0; i < zIndex; i++)
                {
                    shape.ZOrder(MsoZOrderCmd.msoSendBackward);
                }
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
                    return true;
                }

                // 단일 슬라이드 처리
                if (_slide == null)
                {
                    Console.WriteLine("현재 슬라이드가 없습니다.");
                    return false;
                }

                // 슬라이드 번호 처리
                if (lineNumber.StartsWith("Slide: "))
                {
                    int slideNumber = int.Parse(lineNumber.Replace("Slide: ", ""));
                    if (slideNumber > 0 && slideNumber <= _presentation.Slides.Count)
                    {
                        _slide = _presentation.Slides[slideNumber];
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

                    // 위치와 크기 적용
                    if (styleDict.ContainsKey("left"))
                        shape.Left = float.Parse(styleDict["left"].Replace("px", ""));
                    if (styleDict.ContainsKey("top"))
                        shape.Top = float.Parse(styleDict["top"].Replace("px", ""));
                    if (styleDict.ContainsKey("width"))
                        shape.Width = float.Parse(styleDict["width"].Replace("px", ""));
                    if (styleDict.ContainsKey("height"))
                        shape.Height = float.Parse(styleDict["height"].Replace("px", ""));

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

                            // 폰트 크기 적용
                            if (styleDict.ContainsKey("font-size"))
                            {
                                var fontSize = styleDict["font-size"].Replace("pt", "").Trim();
                                if (float.TryParse(fontSize, out float size))
                                {
                                    shape.TextFrame.TextRange.Font.Size = size;
                                }
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
