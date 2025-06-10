using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using WordApp = Microsoft.Office.Interop.Word.Application;
using System.Diagnostics;
using System.IO;
using HtmlAgilityPack;

namespace overlay_gpt
{
    public class WordContextWriter : IContextWriter
    {
        private WordApp? _wordApp;
        private Document? _document;
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
            Guid clsid;
            CLSIDFromProgID(progID, out clsid);
            object obj;
            GetActiveObject(ref clsid, IntPtr.Zero, out obj);
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

                try
                {
                    Console.WriteLine("기존 Word 애플리케이션 찾기 시도...");
                    _wordApp = (WordApp)GetActiveObject("Word.Application");
                    
                    if(_wordApp != null)
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
                    Console.WriteLine($"Word 애플리케이션이 실행 중이지 않습니다: {ex.Message}");
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
            try
            {
                if (_document == null)
                {
                    Console.WriteLine("문서가 열려있지 않습니다.");
                    return false;
                }

                Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] HTML 붙여넣기 시작");
                Console.WriteLine($"입력 HTML 길이: {htmlText.Length} 문자");
                Console.WriteLine($"입력 HTML 내용 샘플 (처음 100자): {htmlText.Substring(0, Math.Min(100, htmlText.Length))}");
                Console.WriteLine($"입력 HTML 내용 샘플 (마지막 100자): {htmlText.Substring(Math.Max(0, htmlText.Length - 100))}");

                // HTML 컨텐츠를 완전한 HTML 문서로 감싸기
                string fullHtml = $@"<html>
<head>
    <meta charset='utf-8'>
    <style>
        body {{ font-family: Arial, sans-serif; font-size: 11pt; }}
        div {{ white-space: pre-wrap; word-wrap: break-word; font-size: 11pt; }}
    </style>
</head>
<body>
    <div>{ProcessImagesInHtml(htmlText)}</div>
</body>
</html>";

                Console.WriteLine($"전체 HTML 길이: {fullHtml.Length} 문자");
                
                // 현재 선택 영역에 HTML 삽입
                Console.WriteLine("Word 문서에 HTML 삽입 시도...");
                Console.WriteLine($"현재 선택 영역 시작 위치: {_document.Application.Selection.Start}");
                Console.WriteLine($"현재 선택 영역 끝 위치: {_document.Application.Selection.End}");

                // 선택 영역 설정
                try
                {
                    // lineNumber에서 시작과 끝 위치 추출
                    string[] positions = lineNumber.Replace("시작: ", "").Replace("끝: ", "").Split(',');
                    if (positions.Length == 2)
                    {
                        int start = int.Parse(positions[0].Trim());
                        int end = int.Parse(positions[1].Trim());
                        
                        Console.WriteLine($"설정할 선택 영역 - 시작: {start}, 끝: {end}");
                        _document.Application.Selection.SetRange(start, end);
                        
                        Console.WriteLine($"선택 영역 설정 후 - 시작: {_document.Application.Selection.Start}, 끝: {_document.Application.Selection.End}");
                        Console.WriteLine($"선택된 텍스트: '{_document.Application.Selection.Text}'");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"선택 영역 설정 중 오류: {ex.Message}");
                }

                // 선택된 범위의 텍스트 삭제
                if (_document.Application.Selection.Start != _document.Application.Selection.End)
                {
                    Console.WriteLine("선택된 범위의 텍스트 삭제 중...");
                    Console.WriteLine($"삭제 전 선택 영역 텍스트: '{_document.Application.Selection.Text}'");
                    Console.WriteLine($"삭제 전 선택 영역 길이: {_document.Application.Selection.Text.Length}");
                    
                    _document.Application.Selection.Delete();
                    
                    Console.WriteLine("선택된 범위의 텍스트 삭제 완료");
                    Console.WriteLine($"삭제 후 선택 영역 시작 위치: {_document.Application.Selection.Start}");
                    Console.WriteLine($"삭제 후 선택 영역 끝 위치: {_document.Application.Selection.End}");
                    Console.WriteLine($"삭제 후 선택 영역 텍스트: '{_document.Application.Selection.Text}'");
                    Console.WriteLine($"삭제 후 선택 영역 길이: {_document.Application.Selection.Text.Length}");
                }
                else
                {
                    Console.WriteLine("선택된 범위가 없습니다. (시작 위치와 끝 위치가 동일)");
                }

                // 임시 파일로 HTML 저장
                string tempFile = Path.GetTempFileName() + ".html";
                File.WriteAllText(tempFile, fullHtml, System.Text.Encoding.UTF8);
                Console.WriteLine($"임시 HTML 파일 생성: {tempFile}");

                try
                {
                    // HTML 파일을 Word에 삽입
                    _document.Application.Selection.InsertFile(tempFile, "", false, false, false);
                    Console.WriteLine("HTML 파일 삽입 완료");

                    // 임시 파일 삭제
                    File.Delete(tempFile);
                    Console.WriteLine("임시 파일 삭제 완료");

                    Console.WriteLine($"삽입 후 선택 영역 시작 위치: {_document.Application.Selection.Start}");
                    Console.WriteLine($"삽입 후 선택 영역 끝 위치: {_document.Application.Selection.End}");
                    Console.WriteLine($"삽입 후 선택 영역 텍스트 길이: {_document.Application.Selection.Text.Length}");
                    Console.WriteLine($"삽입 후 선택 영역 텍스트 샘플: {_document.Application.Selection.Text.Substring(0, Math.Min(100, _document.Application.Selection.Text.Length))}");

                    return true;
                }
                catch (Exception insertEx)
                {
                    Console.WriteLine($"HTML 파일 삽입 실패: {insertEx.Message}");
                    if (File.Exists(tempFile))
                    {
                        File.Delete(tempFile);
                    }
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"텍스트 적용 오류: {ex.Message}");
                Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"내부 예외: {ex.InnerException.Message}");
                    Console.WriteLine($"내부 예외 스택 트레이스: {ex.InnerException.StackTrace}");
                }
                return false;
            }
        }

        private string CreateCFHtmlHeader(string htmlContent)
        {
            Console.WriteLine("CF_HTML 헤더 생성 시작");
            
            // CF_HTML 헤더 형식 생성
            string header = "Version:0.9\r\n";
            header += "StartHTML:0000000000\r\n";
            header += "EndHTML:0000000000\r\n";
            header += "StartFragment:0000000000\r\n";
            header += "EndFragment:0000000000\r\n";
            header += "StartSelection:0000000000\r\n";
            header += "EndSelection:0000000000\r\n";
            header += "SourceURL:file:///C:/temp.html\r\n";

            // HTML 프래그먼트 시작/끝 위치 계산
            int startHtml = header.Length;
            int startFragment = startHtml + htmlContent.IndexOf("<body>") + 6;
            int endFragment = startHtml + htmlContent.IndexOf("</body>");
            int endHtml = startHtml + htmlContent.Length;

            Console.WriteLine($"헤더 길이: {startHtml}");
            Console.WriteLine($"StartFragment 위치: {startFragment}");
            Console.WriteLine($"EndFragment 위치: {endFragment}");
            Console.WriteLine($"EndHTML 위치: {endHtml}");
            Console.WriteLine($"<body> 태그 위치: {htmlContent.IndexOf("<body>")}");
            Console.WriteLine($"</body> 태그 위치: {htmlContent.IndexOf("</body>")}");

            // 헤더의 위치 정보 업데이트
            header = header.Replace("StartHTML:0000000000", $"StartHTML:{startHtml:D10}");
            header = header.Replace("EndHTML:0000000000", $"EndHTML:{endHtml:D10}");
            header = header.Replace("StartFragment:0000000000", $"StartFragment:{startFragment:D10}");
            header = header.Replace("EndFragment:0000000000", $"EndFragment:{endFragment:D10}");
            header = header.Replace("StartSelection:0000000000", $"StartSelection:{startFragment:D10}");
            header = header.Replace("EndSelection:0000000000", $"EndSelection:{endFragment:D10}");

            Console.WriteLine("CF_HTML 헤더 생성 완료");
            return header;
        }

        public (ulong? FileId, uint? VolumeId, string FileType, string FileName, string FilePath) GetFileInfo()
        {
            try
            {
                if (_document == null)
                {
                    return (null, null, "Word", string.Empty, string.Empty);
                }

                string filePath = _document.FullName;
                string fileName = _document.Name;

                return (
                    null, // FileId는 필요시 구현
                    null, // VolumeId는 필요시 구현
                    "Word",
                    fileName,
                    filePath
                );
            }
            catch (Exception ex)
            {
                Console.WriteLine($"파일 정보 가져오기 오류: {ex.Message}");
                return (null, null, "Word", string.Empty, string.Empty);
            }
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

        private string ProcessImagesInHtml(string html)
        {
            try
            {
                var doc = new HtmlDocument();
                doc.LoadHtml(html);

                // 모든 이미지 태그 찾기
                var images = doc.DocumentNode.SelectNodes("//img");
                if (images != null)
                {
                    foreach (var img in images)
                    {
                        var src = img.GetAttributeValue("src", "");
                        if (!string.IsNullOrEmpty(src))
                        {
                            // 이미지 파일 경로 생성
                            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, src);
                            if (File.Exists(imagePath))
                            {
                                // 절대 경로로 변환
                                string absolutePath = Path.GetFullPath(imagePath);
                                img.SetAttributeValue("src", absolutePath);
                                Console.WriteLine($"이미지 경로 변환: {src} -> {absolutePath}");
                            }
                            else
                            {
                                Console.WriteLine($"이미지 파일을 찾을 수 없습니다: {imagePath}");
                            }
                        }
                    }
                }

                return doc.DocumentNode.InnerHtml;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"이미지 처리 중 오류 발생: {ex.Message}");
                return html;
            }
        }
    }
}
