using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;            // System.Windows.Forms.dll 참조 필요
using HtmlAgilityPack;                // HtmlAgilityPack.dll 참조 필요
using WordApp = Microsoft.Office.Interop.Word.Application;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;

namespace overlay_gpt
{
    public class WordContextWriter : IContextWriter, IDisposable
    {
        private WordApp? _wordApp;
        private Word.Document? _document;
        private bool _isTargetProg;

        public bool IsTargetProg
        {
            get => _isTargetProg;
            set => _isTargetProg = value;
        }

        public WordContextWriter(bool isTargetProg = false)
        {
            _isTargetProg = isTargetProg;
        }

        [DllImport("oleaut32.dll")]
        private static extern int GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        [DllImport("ole32.dll")]
        private static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);

        private static object GetActiveObject(string progID)
        {
            CLSIDFromProgID(progID, out Guid clsid);
            GetActiveObject(ref clsid, IntPtr.Zero, out object obj);
            return obj;
        }

        public bool OpenFile(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"파일이 존재하지 않습니다: {filePath}");
                    return false;
                }

                Console.WriteLine("기존 Word 애플리케이션 찾기 시도...");
                _wordApp = (WordApp)GetActiveObject("Word.Application");
                if (_wordApp != null)
                {
                    Console.WriteLine("기존 Word 애플리케이션 찾음");
                    _document = _wordApp.Documents.Open(filePath);
                    return true;
                }
                else
                {
                    Console.WriteLine("기존 Word 애플리케이션을 찾을 수 없습니다.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"파일 열기 오류: {ex.Message}");
                return false;
            }
        }

        public bool ApplyTextWithStyle(string htmlText, string lineNumber)
        {
            if (_document == null)
            {
                Console.WriteLine("문서가 열려있지 않습니다.");
                return false;
            }

            try
            {
                // 1) HTML 전체 문서 래핑
                string bodyHtml = ProcessImagesInHtml(htmlText);
                string fullHtml = $@"<html>
<head>
  <meta charset='utf-8'>
  <style>
    body {{ font-family: Arial, sans-serif; font-size: 11pt; }}
    div {{ white-space: pre-wrap; word-wrap: break-word; font-size: 11pt; }}
    m:oMath {{ font-family: 'Cambria Math', serif; }}
    m:oMathPara {{ margin: 0; padding: 0; }}
  </style>
</head>
<body>
  <div>{bodyHtml}</div>
</body>
</html>";

                // 2) CF_HTML 마커 추가
                string fragmentHtml = $"<!--StartFragment-->{fullHtml}<!--EndFragment-->";
                string cfHeader    = CreateCFHtmlHeader(fragmentHtml);
                string cfHtml      = cfHeader + fragmentHtml;

                // 3) 선택 영역 설정 (lineNumber: "시작: XX, 끝: YY")
                try
                {
                    var parts = lineNumber.Replace("시작:", "").Replace("끝:", "").Split(',');
                    if (parts.Length == 2)
                    {
                        int start = int.Parse(parts[0].Trim());
                        int end   = int.Parse(parts[1].Trim());
                        _wordApp.Selection.SetRange(start, end);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"선택 영역 설정 중 오류: {ex.Message}");
                }

                // 4) 기존 선택 텍스트 삭제
                if (_wordApp.Selection.Start != _wordApp.Selection.End)
                {
                    _wordApp.Selection.Delete();
                }

                // 5) 클립보드에 CF_HTML 올리기
                Clipboard.Clear();
                Clipboard.SetText(cfHtml, TextDataFormat.Html);

                // 6) PasteSpecial 로 HTML 붙여넣기 (내부 HTML→OMML 컨버터 사용)
                int selStart = _wordApp.Selection.Start;
                _wordApp.Selection.PasteSpecial(
                    DataType: Word.WdPasteDataType.wdPasteHTML,
                    Placement: Word.WdOLEPlacement.wdInLine,
                    DisplayAsIcon: false
                );
                int selEnd = _wordApp.Selection.End;
                if (selStart != selEnd)
                {
                    Word.Range pastedRange = _document.Range(selStart, selEnd);
                    _document.OMaths.Add(pastedRange);
                    _document.OMaths.BuildUp();
                }
                Console.WriteLine("HTML PasteSpecial 완료");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"텍스트 적용 오류: {ex.Message}");
                return false;
            }
        }

        private string CreateCFHtmlHeader(string htmlFragment)
        {
            // CF_HTML 헤더 틀
            string header = 
                "Version:0.9\r\n" +
                "StartHTML:??????????\r\n" +
                "EndHTML:??????????\r\n" +
                "StartFragment:??????????\r\n" +
                "EndFragment:??????????\r\n" +
                "SourceURL:file:///\r\n";

            // 위치 계산
            int startHTML     = header.Length;
            int startFragment = startHTML + htmlFragment.IndexOf("<!--StartFragment-->", StringComparison.Ordinal) + "<!--StartFragment-->".Length;
            int endFragment   = startHTML + htmlFragment.IndexOf("<!--EndFragment-->", StringComparison.Ordinal);
            int endHTML       = startHTML + htmlFragment.Length;

            // 숫자 보정 (10자리 폭)
            header = header
                .Replace("StartHTML:??????????",      $"StartHTML:{startHTML:D10}")
                .Replace("EndHTML:??????????",        $"EndHTML:{endHTML:D10}")
                .Replace("StartFragment:??????????",  $"StartFragment:{startFragment:D10}")
                .Replace("EndFragment:??????????",    $"EndFragment:{endFragment:D10}");

            return header;
        }

        private string ProcessImagesInHtml(string html)
        {
            try
            {
                Console.WriteLine("HTML 처리 시작...");
                var doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(html);
                Console.WriteLine($"원본 HTML 길이: {html.Length}");

                // conditional comment 수식 처리
                var comments = doc.DocumentNode.SelectNodes("//comment()");
                if (comments != null)
                {
                    Console.WriteLine($"조건부 주석 수식 발견: {comments.Count}개");
                    foreach (var cm in comments)
                    {
                        if (cm.InnerHtml.Contains("if gte msEquation"))
                        {
                            Console.WriteLine("msEquation 수식 발견");
                            int s = cm.InnerHtml.IndexOf("]>") + 2;
                            int e = cm.InnerHtml.IndexOf("<![endif]");
                            if (s > 1 && e > s)
                            {
                                var frag = new HtmlAgilityPack.HtmlDocument();
                                string inner = cm.InnerHtml.Substring(s, e - s);
                                Console.WriteLine($"수식 내용: {inner}");
                                frag.LoadHtml(inner);
                                
                                // 수식 노드에 필요한 네임스페이스 추가
                                var mathNodes = frag.DocumentNode.SelectNodes("//*[local-name()='oMath']");
                                if (mathNodes != null)
                                {
                                    Console.WriteLine($"oMath 노드 발견: {mathNodes.Count}개");
                                    foreach (var m in mathNodes)
                                    {
                                        m.SetAttributeValue("xmlns:m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
                                        m.SetAttributeValue("xmlns:o", "urn:schemas-microsoft-com:office:office");
                                        Console.WriteLine($"oMath 노드 처리됨: {m.OuterHtml}");
                                    }
                                }

                                // 수식 단락에 필요한 네임스페이스 추가
                                var mathParaNodes = frag.DocumentNode.SelectNodes("//*[local-name()='oMathPara']");
                                if (mathParaNodes != null)
                                {
                                    Console.WriteLine($"oMathPara 노드 발견: {mathParaNodes.Count}개");
                                    foreach (var mp in mathParaNodes)
                                    {
                                        mp.SetAttributeValue("xmlns:m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
                                        mp.SetAttributeValue("xmlns:o", "urn:schemas-microsoft-com:office:office");
                                        Console.WriteLine($"oMathPara 노드 처리됨: {mp.OuterHtml}");
                                    }
                                }

                                // 원본 스타일 추출
                                var styleNode = frag.DocumentNode.SelectSingleNode("//span[@style]");
                                string originalStyle = styleNode?.GetAttributeValue("style", "") ?? "";
                                Console.WriteLine($"원본 스타일: {originalStyle}");

                                // 수식 단락을 Word 필드로 처리하면서 원본 스타일 유지
                                Console.WriteLine($"원본 수식 내용: {inner}");
                                Console.WriteLine($"수식 처리 시작: inner 길이 = {inner.Length}");
                                string trimmedInner = inner.Trim();
                                Console.WriteLine($"수식 처리: Trim 후 길이 = {trimmedInner.Length}");
                                string replacedInner = trimmedInner.Replace("<m:oMathPara>", "").Replace("</m:oMathPara>", "");
                                replacedInner = Regex.Replace(replacedInner, "<\\/?(span|b|i)[^>]*>", "");
                                Console.WriteLine($"수식 처리: Replace 후 결과 문자열 길이 = {replacedInner.Length}");
                                Console.WriteLine($"수식 처리: Trim 결과 문자열 = {trimmedInner}");
                                Console.WriteLine($"수식 처리: Replace 후 결과 내용 = {replacedInner}");
                                if(replacedInner.Contains("<m:oMathPara>") || replacedInner.Contains("</m:oMathPara>")) {
                                    Console.WriteLine("경고: Replace 후에도 <m:oMathPara> 태그가 남아있습니다.");
                                }
                                string processedMath = "<m:oMathPara xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math' xmlns:o='urn:schemas-microsoft-com:office:office'>" +
                                                         replacedInner +
                                                         "</m:oMathPara>";
                                Console.WriteLine($"최종 처리된 수식: {processedMath}");
                                cm.ParentNode.InnerHtml = cm.ParentNode.InnerHtml.Replace(cm.InnerHtml, processedMath);
                            }
                        }
                    }
                }

                // 일반 MathML 처리
                var mathAll = doc.DocumentNode.SelectNodes("//*[local-name()='oMath']");
                if (mathAll != null)
                {
                    Console.WriteLine($"일반 MathML 노드 발견: {mathAll.Count}개");
                    foreach (var m in mathAll)
                    {
                        m.SetAttributeValue("xmlns:m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
                        m.SetAttributeValue("xmlns:o", "urn:schemas-microsoft-com:office:office");
                        Console.WriteLine($"MathML 노드 처리됨: {m.OuterHtml}");
                    }
                }

                // <img> 경로 절대화
                var imgs = doc.DocumentNode.SelectNodes("//img");
                if (imgs != null)
                {
                    Console.WriteLine($"이미지 태그 발견: {imgs.Count}개");
                    foreach (var img in imgs)
                    {
                        var src = img.GetAttributeValue("src", "");
                        if (!string.IsNullOrEmpty(src) && File.Exists(src))
                        {
                            img.SetAttributeValue("src", Path.GetFullPath(src));
                            Console.WriteLine($"이미지 경로 절대화: {src} -> {Path.GetFullPath(src)}");
                        }
                    }
                }

                string result = doc.DocumentNode.InnerHtml;
                Console.WriteLine($"처리된 HTML 길이: {result.Length}");
                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"HTML 처리 중 오류 발생: {ex.Message}");
                return html;
            }
        }

        public (ulong? FileId, uint? VolumeId, string FileType, string FileName, string FilePath) GetFileInfo()
        {
            if (_document == null)
                return (null, null, "Word", "", "");
            return (null, null, "Word", _document.Name, _document.FullName);
        }

        public void Dispose()
        {
            if (_document != null)
            {
                Marshal.ReleaseComObject(_document);
                _document = null;
            }
            if (_wordApp != null)
            {
                Marshal.ReleaseComObject(_wordApp);
                _wordApp = null;
            }
        }
    }
}
