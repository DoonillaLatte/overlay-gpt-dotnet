using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Diagnostics;
using HtmlAgilityPack;

namespace overlay_gpt
{
    public class PPTContextWriter : IContextWriter
    {
        private Application? _pptApp;
        private Presentation? _presentation;
        private Slide? _slide;

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

        private int ConvertColorToRGB(int bgrColor)
        {
            int r = bgrColor & 0xFF;
            int g = (bgrColor >> 8) & 0xFF;
            int b = (bgrColor >> 16) & 0xFF;
            return (r << 16) | (g << 8) | b;
        }

        public bool OpenFile(string filePath)
        {
            try
            {
                Console.WriteLine("기존 PowerPoint 프로세스 확인 중...");
                try
                {
                    _pptApp = (Application)GetActiveObject("PowerPoint.Application");
                    Console.WriteLine("기존 PowerPoint 프로세스 발견");

                    // 이미 열려있는 프레젠테이션 확인
                    foreach (Presentation pres in _pptApp.Presentations)
                    {
                        try
                        {
                            if (pres.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("파일이 이미 열려있습니다.");
                                _presentation = pres;
                                _slide = _pptApp.ActiveWindow?.View?.Slide;
                                return true;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"프레젠테이션 확인 중 오류 발생: {ex.Message}");
                            continue;
                        }
                    }
                }
                catch
                {
                    Console.WriteLine("새로운 PowerPoint COM 객체 생성 시도...");
                    _pptApp = new Application();
                    _pptApp.Visible = MsoTriState.msoFalse; // 백그라운드에서 실행
                    Console.WriteLine("새로운 PowerPoint COM 객체 생성 성공");
                }

                Console.WriteLine($"파일 열기 시도: {filePath}");
                _presentation = _pptApp.Presentations.Open(filePath);
                _slide = _pptApp.ActiveWindow?.View?.Slide;
                Console.WriteLine("파일 열기 성공");

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"PowerPoint 파일 열기 오류 발생: {ex.Message}");
                Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                
                // 오류 발생 시 COM 객체 정리
                if (_slide != null)
                {
                    try { Marshal.ReleaseComObject(_slide); } catch { }
                    _slide = null;
                }
                if (_presentation != null)
                {
                    try { Marshal.ReleaseComObject(_presentation); } catch { }
                    _presentation = null;
                }
                if (_pptApp != null)
                {
                    try { Marshal.ReleaseComObject(_pptApp); } catch { }
                    _pptApp = null;
                }
                
                return false;
            }
        }

        public bool ApplyTextWithStyle(string text, string lineNumber)
        {
            try
            {
                if (_pptApp == null || _slide == null)
                {
                    Console.WriteLine("PowerPoint 애플리케이션이 초기화되지 않았습니다.");
                    return false;
                }

                Console.WriteLine($"텍스트 적용 시작 - 라인 번호: {lineNumber}");
                Console.WriteLine($"적용할 텍스트: {text}");

                // 라인 번호 파싱 (예: "Slide 1")
                var slideNumber = int.Parse(lineNumber.Replace("Slide ", ""));
                Console.WriteLine($"슬라이드 번호: {slideNumber}");

                // HTML 태그 처리
                Console.WriteLine("HTML 파싱 시작...");
                var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.LoadHtml(text);
                Console.WriteLine($"HTML 노드 수: {htmlDoc.DocumentNode.ChildNodes.Count}");

                // 기존 도형 삭제
                foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in _slide.Shapes)
                {
                    try
                    {
                        shape.Delete();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"도형 삭제 중 오류 발생: {ex.Message}");
                    }
                }

                // HTML 노드 처리
                foreach (var node in htmlDoc.DocumentNode.ChildNodes)
                {
                    try
                    {
                        Console.WriteLine($"노드 처리 시작 - 타입: {node.NodeType}, 이름: {node.Name}");
                        
                        if (node.NodeType == HtmlAgilityPack.HtmlNodeType.Element)
                        {
                            var shape = _slide.Shapes.AddShape(
                                MsoAutoShapeType.msoShapeRectangle,
                                0, 0, 100, 100);

                            // 스타일 속성 파싱
                            var style = node.GetAttributeValue("style", "");
                            Console.WriteLine($"스타일 속성: {style}");
                            
                            var styleAttributes = style.Split(';')
                                .Select(s => s.Trim().Split(':'))
                                .Where(p => p.Length == 2)
                                .ToDictionary(p => p[0].Trim(), p => p[1].Trim());

                            // 위치와 크기 설정
                            if (styleAttributes.TryGetValue("left", out var left))
                                shape.Left = float.Parse(left.Replace("px", ""));
                            if (styleAttributes.TryGetValue("top", out var top))
                                shape.Top = float.Parse(top.Replace("px", ""));
                            if (styleAttributes.TryGetValue("width", out var width))
                                shape.Width = float.Parse(width.Replace("px", ""));
                            if (styleAttributes.TryGetValue("height", out var height))
                                shape.Height = float.Parse(height.Replace("px", ""));

                            // 텍스트 설정
                            if (node.InnerText != null)
                            {
                                shape.TextFrame.TextRange.Text = node.InnerText;
                                var textRange = shape.TextFrame.TextRange;

                                // 텍스트 스타일 설정
                                if (styleAttributes.TryGetValue("font-family", out var fontFamily))
                                    textRange.Font.Name = fontFamily.Trim('\'');
                                if (styleAttributes.TryGetValue("font-size", out var fontSize))
                                    textRange.Font.Size = float.Parse(fontSize.Replace("pt", ""));
                                if (styleAttributes.TryGetValue("color", out var color))
                                {
                                    if (color.StartsWith("#"))
                                    {
                                        var rgb = int.Parse(color.Substring(1), System.Globalization.NumberStyles.HexNumber);
                                        textRange.Font.Color.RGB = ConvertColorToRGB(rgb);
                                    }
                                }

                                // 텍스트 정렬
                                if (styleAttributes.TryGetValue("text-align", out var textAlign))
                                {
                                    switch (textAlign)
                                    {
                                        case "center":
                                            textRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                                            break;
                                        case "right":
                                            textRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignRight;
                                            break;
                                        case "justify":
                                            textRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignJustify;
                                            break;
                                        default:
                                            textRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
                                            break;
                                    }
                                }

                                // 배경색 설정
                                if (styleAttributes.TryGetValue("background-color", out var bgColor))
                                {
                                    if (bgColor.StartsWith("#"))
                                    {
                                        var rgb = int.Parse(bgColor.Substring(1), System.Globalization.NumberStyles.HexNumber);
                                        shape.Fill.ForeColor.RGB = ConvertColorToRGB(rgb);
                                    }
                                }

                                // 투명도 설정
                                if (styleAttributes.TryGetValue("opacity", out var opacity))
                                {
                                    shape.Fill.Transparency = (1 - float.Parse(opacity)) * 100;
                                }

                                // 테두리 설정
                                if (styleAttributes.TryGetValue("border", out var border))
                                {
                                    var borderParts = border.Split(' ');
                                    if (borderParts.Length >= 3)
                                    {
                                        shape.Line.Weight = float.Parse(borderParts[0].Replace("px", ""));
                                        shape.Line.ForeColor.RGB = int.Parse(borderParts[2].Replace("#", ""), System.Globalization.NumberStyles.HexNumber);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"노드 처리 중 오류 발생: {ex.Message}");
                        Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                    }
                }

                Console.WriteLine("텍스트 적용 완료");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"텍스트 적용 중 오류 발생: {ex.Message}");
                Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                return false;
            }
        }

        public (ulong? FileId, uint? VolumeId, string FileType, string FileName, string FilePath) GetFileInfo()
        {
            Application? tempPptApp = null;
            Presentation? tempPresentation = null;
            
            try
            {
                Console.WriteLine("PowerPoint COM 객체 가져오기 시도...");
                tempPptApp = (Application)GetActiveObject("PowerPoint.Application");
                Console.WriteLine("PowerPoint COM 객체 가져오기 성공");

                Console.WriteLine("활성 프레젠테이션 가져오기 시도...");
                tempPresentation = tempPptApp.ActivePresentation;
                
                if (tempPresentation == null)
                {
                    Console.WriteLine("활성 프레젠테이션을 찾을 수 없습니다.");
                    return (null, null, "PowerPoint", string.Empty, string.Empty);
                }

                string filePath = tempPresentation.FullName;
                string fileName = tempPresentation.Name;
                
                Console.WriteLine($"PowerPoint 프레젠테이션 정보:");
                Console.WriteLine($"- 파일 경로: {filePath}");
                Console.WriteLine($"- 파일 이름: {fileName}");
                
                if (string.IsNullOrEmpty(filePath))
                {
                    Console.WriteLine("파일 경로가 비어있습니다.");
                    return (null, null, "PowerPoint", fileName, string.Empty);
                }
                
                return (null, null, "PowerPoint", fileName, filePath);
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

        public void Dispose()
        {
            if (_slide != null)
            {
                try { Marshal.ReleaseComObject(_slide); } catch { }
                _slide = null;
            }
            if (_presentation != null)
            {
                try { Marshal.ReleaseComObject(_presentation); } catch { }
                _presentation = null;
            }
            if (_pptApp != null)
            {
                try { Marshal.ReleaseComObject(_pptApp); } catch { }
                _pptApp = null;
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
