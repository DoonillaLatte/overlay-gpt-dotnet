using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using HtmlAgilityPack;

namespace overlay_gpt
{
    public class WordContextWriter : IContextWriter
    {
        private Application? _wordApp;
        private Document? _document;

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

        private WdColorIndex GetHighlightColorIndexFromRGB(int rgbColor)
        {
            // RGB 색상에 따른 WdColorIndex 매핑
            switch (rgbColor)
            {
                case 0xFFFF00: return WdColorIndex.wdYellow;  // 노랑
                case 0x00FF00: return WdColorIndex.wdBrightGreen;  // 밝은 초록
                case 0x00FFFF: return WdColorIndex.wdTurquoise;  // 청록
                case 0xFF00FF: return WdColorIndex.wdPink;  // 분홍     
                case 0x0000FF: return WdColorIndex.wdBlue;  // 파랑
                case 0xFF0000: return WdColorIndex.wdRed;  // 빨강
                case 0x000080: return WdColorIndex.wdDarkBlue;  // 진한 파랑
                case 0x008080: return WdColorIndex.wdTeal;  // 청녹
                case 0x008000: return WdColorIndex.wdGreen;  // 초록
                case 0x800080: return WdColorIndex.wdViolet;  // 보라
                case 0x800000: return WdColorIndex.wdDarkRed;  // 진한 빨강
                case 0x808000: return WdColorIndex.wdDarkYellow;  // 진한 노랑
                case 0x808080: return WdColorIndex.wdGray50;  // 회색
                case 0xC0C0C0: return WdColorIndex.wdGray25;  // 연한 회색
                default: return WdColorIndex.wdNoHighlight;  // 기본값 (하이라이트 없음)
            }
        }

        public bool OpenFile(string filePath)
        {
            try
            {
                Console.WriteLine("기존 Word 프로세스 확인 중...");
                try
                {
                    _wordApp = (Application)GetActiveObject("Word.Application");
                    Console.WriteLine("기존 Word 프로세스 발견");

                    // 이미 열려있는 문서 확인
                    foreach (Document doc in _wordApp.Documents)
                    {
                        try
                        {
                            if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("파일이 이미 열려있습니다.");
                                _document = doc;
                                return true;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"문서 확인 중 오류 발생: {ex.Message}");
                            continue;
                        }
                    }
                }
                catch
                {
                    Console.WriteLine("새로운 Word COM 객체 생성 시도...");
                    _wordApp = new Application();
                    _wordApp.Visible = false; // 백그라운드에서 실행
                    Console.WriteLine("새로운 Word COM 객체 생성 성공");
                }

                Console.WriteLine($"파일 열기 시도: {filePath}");
                _document = _wordApp.Documents.Open(filePath);
                Console.WriteLine("파일 열기 성공");

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Word 파일 열기 오류 발생: {ex.Message}");
                Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                
                // 오류 발생 시 COM 객체 정리
                if (_document != null)
                {
                    try { Marshal.ReleaseComObject(_document); } catch { }
                    _document = null;
                }
                if (_wordApp != null)
                {
                    try { Marshal.ReleaseComObject(_wordApp); } catch { }
                    _wordApp = null;
                }
                
                return false;
            }
        }

        public bool ApplyTextWithStyle(string text, string lineNumber)
        {
            try
            {
                if (_wordApp == null || _document == null)
                {
                    Console.WriteLine("Word 애플리케이션이 초기화되지 않았습니다.");
                    return false;
                }

                Console.WriteLine($"텍스트 적용 시작 - 라인 번호: {lineNumber}");
                Console.WriteLine($"적용할 텍스트: {text}");

                // 라인 번호 파싱 (예: "1-1")
                var lineNumbers = lineNumber.Split('-');
                if (lineNumbers.Length != 2)
                {
                    Console.WriteLine("잘못된 라인 번호 형식입니다.");
                    return false;
                }

                int startLine = int.Parse(lineNumbers[0]);
                int endLine = int.Parse(lineNumbers[1]);
                Console.WriteLine($"시작 라인: {startLine}, 종료 라인: {endLine}");

                // 해당 라인으로 이동
                Microsoft.Office.Interop.Word.Range? wordRange = null;
                try
                {
                    Console.WriteLine("라인 범위 설정 시도...");
                    
                    // 시작 라인으로 이동
                    var startRange = _document.GoTo(WdGoToItem.wdGoToLine, WdGoToDirection.wdGoToAbsolute, startLine);
                    Console.WriteLine($"시작 라인 위치: {startRange.Start}");
                    
                    // 종료 라인으로 이동
                    var endRange = _document.GoTo(WdGoToItem.wdGoToLine, WdGoToDirection.wdGoToAbsolute, endLine);
                    Console.WriteLine($"종료 라인 위치: {endRange.End}");
                    
                    // 라인 범위 설정
                    wordRange = _document.Range(startRange.Start, endRange.End);
                    Console.WriteLine("라인 범위 설정 성공");

                    // 기존 텍스트 지우기
                    Console.WriteLine("기존 텍스트 지우기");
                    
                    // 각 라인을 순회하며 텍스트 지우기
                    for (int line = startLine; line <= endLine; line++)
                    {
                        var lineRange = _document.GoTo(WdGoToItem.wdGoToLine, WdGoToDirection.wdGoToAbsolute, line);
                        var nextLineRange = _document.GoTo(WdGoToItem.wdGoToLine, WdGoToDirection.wdGoToAbsolute, line + 1);
                        
                        // 현재 라인의 시작부터 다음 라인의 시작까지의 범위를 지움
                        var fullLineRange = _document.Range(lineRange.Start, nextLineRange.Start);
                        Console.WriteLine($"라인 {line} 텍스트 지우기 (시작: {lineRange.Start}, 끝: {nextLineRange.Start})");
                        fullLineRange.Delete();
                    }
                    
                    Console.WriteLine("기존 텍스트 삭제 완료");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"라인 범위 설정 중 오류 발생: {ex.Message}");
                    Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                    return false;
                }

                if (wordRange is null)
                {
                    Console.WriteLine("라인 범위를 설정할 수 없습니다.");
                    return false;
                }

                // HTML 태그 처리
                Console.WriteLine("HTML 파싱 시작...");
                var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.LoadHtml(text);
                Console.WriteLine($"HTML 노드 수: {htmlDoc.DocumentNode.ChildNodes.Count}");

                // 텍스트와 스타일 적용
                int nodeIndex = 0;
                foreach (var node in htmlDoc.DocumentNode.ChildNodes)
                {
                    try
                    {
                        Console.WriteLine($"노드 {nodeIndex} 처리 시작 - 타입: {node.NodeType}, 이름: {node.Name}");
                        
                        // 현재 노드의 범위 생성
                        var currentNodeRange = _document.Range(wordRange.Start, wordRange.Start);
                        
                        if (node.NodeType == HtmlAgilityPack.HtmlNodeType.Text)
                        {
                            Console.WriteLine($"텍스트 노드 처리 - 내용: {node.InnerText}");
                            currentNodeRange.Text = node.InnerText;
                        }
                        else
                        {
                            Console.WriteLine($"HTML 태그 노드 처리 - 태그: {node.Name}");
                            var style = node.GetAttributeValue("style", "");
                            Console.WriteLine($"스타일 속성: {style}");
                            
                            // 먼저 텍스트를 삽입
                            currentNodeRange.Text = node.InnerText;
                            
                            // 텍스트 삽입 후 범위 다시 설정
                            currentNodeRange = _document.Range(wordRange.Start, wordRange.Start + node.InnerText.Length);
                            
                            var font = currentNodeRange.Font;
                            Console.WriteLine("폰트 객체 가져오기 성공");

                            // 스타일 속성 파싱
                            var styleAttributes = style.Split(';')
                                .Select(s => s.Trim().Split(':'))
                                .Where(p => p.Length == 2)
                                .ToDictionary(p => p[0].Trim(), p => p[1].Trim());

                            Console.WriteLine($"파싱된 스타일 속성 수: {styleAttributes.Count}");

                            // 폰트 패밀리
                            if (styleAttributes.TryGetValue("font-family", out var fontFamily))
                            {
                                Console.WriteLine($"폰트 패밀리 설정: {fontFamily}");
                                font.Name = fontFamily.Trim('\'');
                            }

                            // 폰트 크기
                            if (styleAttributes.TryGetValue("font-size", out var fontSize))
                            {
                                if (fontSize.EndsWith("pt"))
                                {
                                    Console.WriteLine($"폰트 크기 설정: {fontSize}");
                                    font.Size = float.Parse(fontSize.Replace("pt", ""));
                                }
                            }

                            // 색상
                            if (styleAttributes.TryGetValue("color", out var color))
                            {
                                if (color.StartsWith("#"))
                                {
                                    Console.WriteLine($"텍스트 색상 설정: {color}");
                                    var rgb = int.Parse(color.Substring(1), System.Globalization.NumberStyles.HexNumber);
                                    font.Color = (WdColor)rgb;
                                }
                            }

                            // 배경색
                            if (styleAttributes.TryGetValue("background-color", out var bgColor))
                            {
                                if (bgColor.StartsWith("#"))
                                {
                                    Console.WriteLine($"배경색 설정: {bgColor}");
                                    var rgb = int.Parse(bgColor.Substring(1), System.Globalization.NumberStyles.HexNumber);
                                    currentNodeRange.HighlightColorIndex = GetHighlightColorIndexFromRGB(rgb);
                                }
                            }

                            // 굵게
                            if (node.Name == "b" || node.Name == "strong")
                            {
                                Console.WriteLine("굵게 스타일 적용");
                                font.Bold = -1;
                            }

                            // 기울임
                            if (node.Name == "i" || node.Name == "em")
                            {
                                Console.WriteLine("기울임 스타일 적용");
                                font.Italic = -1;
                            }

                            // 밑줄
                            if (node.Name == "u")
                            {
                                Console.WriteLine("밑줄 스타일 적용");
                                font.Underline = WdUnderline.wdUnderlineSingle;
                            }

                            // 취소선
                            if (node.Name == "s" || node.Name == "strike")
                            {
                                Console.WriteLine("취소선 스타일 적용");
                                font.StrikeThrough = -1;
                            }
                        }
                        
                        // 다음 노드를 위한 범위 업데이트
                        wordRange.Start = currentNodeRange.End;
                        
                        Console.WriteLine($"노드 {nodeIndex} 처리 완료");
                        nodeIndex++;
                        
                        // COM 객체 정리
                        if (currentNodeRange != null)
                        {
                            try { Marshal.ReleaseComObject(currentNodeRange); } catch { }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"노드 {nodeIndex} 처리 중 오류 발생: {ex.Message}");
                        Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                        continue;
                    }
                }

                // 줄바꿈 추가
                try
                {
                    Console.WriteLine("줄바꿈 추가");
                    wordRange.InsertAfter("\r");
                    Console.WriteLine("줄바꿈 추가 완료");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"줄바꿈 추가 중 오류 발생: {ex.Message}");
                }

                // COM 객체 정리
                if (wordRange != null)
                {
                    try { Marshal.ReleaseComObject(wordRange); } catch { }
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
                
                return (null, null, "Word", fileName, filePath);
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

        public void Dispose()
        {
            if (_document != null)
            {
                try { Marshal.ReleaseComObject(_document); } catch { }
                _document = null;
            }
            if (_wordApp != null)
            {
                try { Marshal.ReleaseComObject(_wordApp); } catch { }
                _wordApp = null;
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
} 