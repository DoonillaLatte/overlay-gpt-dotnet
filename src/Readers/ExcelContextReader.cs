using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using System.IO;
using System.Text;
using System.Diagnostics;

namespace overlay_gpt
{
    public class ExcelContextReader : BaseContextReader
    {
        private Microsoft.Office.Interop.Excel.Application? _excelApp;
        private Workbook? _workbook;
        private bool _isTargetProg;
        private string? _filePath;

        public ExcelContextReader(bool isTargetProg = false, string filePath = "")
        {
            Console.WriteLine($"ExcelContextReader 생성 시도 - isTargetProg: {isTargetProg}");
            _isTargetProg = isTargetProg;
            _filePath = filePath;
        }

        private Dictionary<string, string> ExtractStylesFromHtml(string htmlContent)
        {
            var styles = new Dictionary<string, string>();
            
            // style 태그 내용 추출
            var styleTagPattern = @"<style[^>]*>(.*?)</style>";
            var styleTagMatch = Regex.Match(htmlContent, styleTagPattern, RegexOptions.Singleline);
            
            if (styleTagMatch.Success)
            {
                string styleTagContent = styleTagMatch.Groups[1].Value;
                var stylePattern = @"\.(xl\d+)\s*\{([^}]+)\}";
                var matches = Regex.Matches(styleTagContent, stylePattern);

                foreach (Match match in matches)
                {
                    string className = match.Groups[1].Value;
                    string styleContent = match.Groups[2].Value
                        .Replace("\r", "")
                        .Replace("\n", "")
                        .Replace("\t", "")
                        .Trim();
                    styles[className] = styleContent;
                }
            }

            return styles;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        private bool IsExcelProcessActive()
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

                // Excel 프로세스 이름 확인 (EXCEL.EXE)
                return foregroundProcess.ProcessName.Equals("EXCEL", StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"프로세스 확인 중 오류 발생: {ex.Message}");
                return false;
            }
        }

        public override (string SelectedText, Dictionary<string, object> StyleAttributes, string LineNumber) GetSelectedTextWithStyle(bool readAllContent = false)
        {
            try
            {
                Console.WriteLine("Excel 데이터 읽기 시작...");

                // 현재 포커스된 프로세스가 Excel인지 확인
                if (!IsExcelProcessActive())
                {
                    Console.WriteLine("현재 포커스된 프로세스가 Excel이 아닙니다.");
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);
                }

                // Excel COM 객체 생성 시도
                try
                {
                    Console.WriteLine("기존 Excel 애플리케이션 찾기 시도...");
                    _excelApp = (Microsoft.Office.Interop.Excel.Application)GetActiveObject("Excel.Application");
                    
                    
                    if (_excelApp != null)
                    {
                        Console.WriteLine("기존 Excel 애플리케이션 찾음");
                        Console.WriteLine($"Excel 버전: {_excelApp.Version}");
                        Console.WriteLine($"활성 워크북 수: {_excelApp.Workbooks.Count}");
                        
                        // 활성 워크북 가져오기
                        _workbook = _excelApp.ActiveWorkbook;
                        if (_workbook == null)
                        {
                            Console.WriteLine("활성 워크북을 찾을 수 없습니다.");
                            return (string.Empty, new Dictionary<string, object>(), string.Empty);
                        }
                        Console.WriteLine($"활성 워크북 이름: {_workbook.Name}");
                    }
                    else 
                    {
                        if (_isTargetProg)
                        {
                            Console.WriteLine("기존 프로세스가 없어, 새 Excel 프로세스를 생성합니다.");
                            try
                            {
                                _excelApp = new Microsoft.Office.Interop.Excel.Application();
                                Console.WriteLine("새 Excel 애플리케이션 생성 성공");
                                _excelApp.Visible = false;  // Excel 창을 안보이게 설정

                                Console.WriteLine($"활성 파일 경로: {_filePath}");

                                if (!string.IsNullOrEmpty(_filePath))
                                {
                                    Console.WriteLine($"파일 열기 시도: {_filePath}");
                                    _workbook = _excelApp.Workbooks.Open(_filePath);
                                    Console.WriteLine("파일 열기 성공");
                                }
                                Console.WriteLine($"활성 워크북 상태: {(_workbook != null ? "존재함" : "없음")}");
                            }
                            catch (Exception createEx)
                            {
                                Console.WriteLine($"새 Excel 애플리케이션 생성 실패: {createEx.Message}");
                                Console.WriteLine($"스택 트레이스: {createEx.StackTrace}");
                                throw;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"기존 Excel 애플리케이션 찾기 실패: {ex.Message}");
                    Console.WriteLine("Excel 애플리케이션이 실행 중이지 않습니다.");
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);
                }

                if (_excelApp == null)
                {
                    Console.WriteLine("Excel 애플리케이션을 생성할 수 없습니다.");
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);
                }

                // 선택된 범위 가져오기
                Microsoft.Office.Interop.Excel.Range? selection = null;
                if (_isTargetProg)
                {
                    Console.WriteLine("전체 범위 선택");
                    Worksheet worksheet = _workbook.ActiveSheet;
                    Console.WriteLine($"현재 워크시트: {worksheet.Name}");
                    
                    selection = worksheet.UsedRange;
                    Console.WriteLine($"선택된 범위: {selection.Address}");
                    Console.WriteLine($"선택된 범위 행 수: {selection.Rows.Count}");
                    Console.WriteLine($"선택된 범위 열 수: {selection.Columns.Count}");
                    
                    try 
                    {
                        // Excel을 일시적으로 보이게 설정
                        bool originalVisible = _excelApp.Visible;
                        _excelApp.Visible = true;
                        
                        // 선택된 범위를 클립보드에 복사
                        Console.WriteLine("클립보드에 복사 시도...");
                        selection.Copy();
                        Console.WriteLine("클립보드 복사 완료");
                        
                        // 원래 상태로 복원
                        _excelApp.Visible = originalVisible;
                        
                        // 클립보드 내용 확인
                        Console.WriteLine("클립보드 형식 확인 중...");
                        Console.WriteLine($"HTML 형식 존재: {Clipboard.ContainsText(TextDataFormat.Html)}");
                        Console.WriteLine($"일반 텍스트 존재: {Clipboard.ContainsText()}");
                        Console.WriteLine($"RTF 형식 존재: {Clipboard.ContainsText(TextDataFormat.Rtf)}");
                        
                        if (Clipboard.ContainsText(TextDataFormat.Html))
                        {
                            string htmlContent = Clipboard.GetText(TextDataFormat.Html);
                            Console.WriteLine($"HTML 데이터 길이: {htmlContent.Length}");
                            Console.WriteLine("HTML 데이터 일부: " + htmlContent.Substring(0, Math.Min(100, htmlContent.Length)));
                            return (htmlContent, new Dictionary<string, object>(), selection.Address);
                        }
                        else
                        {
                            Console.WriteLine("클립보드에 HTML 형식 데이터 없음");
                            return (string.Empty, new Dictionary<string, object>(), string.Empty);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"클립보드 복사 중 오류 발생: {ex.Message}");
                        Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                        return (string.Empty, new Dictionary<string, object>(), string.Empty);
                    }
                }
                else
                {
                    Console.WriteLine("선택된 범위 가져오기");
                    selection = _excelApp.Selection as Microsoft.Office.Interop.Excel.Range;
                    if (selection != null)
                    {
                        selection.Copy();
                    }
                }

                if (selection == null)
                {
                    Console.WriteLine("선택된 범위가 없습니다.");
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);
                }

                // HTML 형식으로 클립보드에서 가져오기
                if (Clipboard.ContainsText(TextDataFormat.Html))
                {
                    string htmlContent = Clipboard.GetText(TextDataFormat.Html);
                    
                    // 원본 HTML 데이터 출력
                    Console.WriteLine("\n=== 원본 HTML 데이터 ===");
                    Console.WriteLine(htmlContent);
                    Console.WriteLine("========================\n");
                    
                    // 스타일 추출
                    var styleAttributes = new Dictionary<string, object>();
                    var styles = ExtractStylesFromHtml(htmlContent);
                    foreach (var style in styles)
                    {
                        styleAttributes[style.Key] = style.Value;
                    }

                    // 테이블 태그만 추출
                    var tablePattern = @"<table[^>]*>.*?</table>";
                    var tableMatch = Regex.Match(htmlContent, tablePattern, RegexOptions.Singleline);
                    if (tableMatch.Success)
                    {
                        htmlContent = tableMatch.Value;
                    }
                    
                    // HtmlAgilityPack을 사용하여 HTML 파싱 및 수정
                    var doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(htmlContent);

                    // class 속성을 가진 모든 노드 찾기
                    var nodes = doc.DocumentNode.SelectNodes("//*[@class]");
                    if (nodes != null)
                    {
                        foreach (var node in nodes)
                        {
                            string className = node.GetAttributeValue("class", "");
                            if (styles.ContainsKey(className))
                            {
                                node.SetAttributeValue("style", styles[className]);
                                node.Attributes.Remove("class");
                            }
                        }
                    }

                    // 수정된 HTML 가져오기
                    htmlContent = doc.DocumentNode.OuterHtml;

                    // 스타일 정보 출력
                    Console.WriteLine("\n=== 추출된 스타일 정보 ===");
                    foreach (var style in styles)
                    {
                        Console.WriteLine($"클래스: {style.Key}");
                        Console.WriteLine($"스타일: {style.Value}");
                        Console.WriteLine("------------------------");
                    }
                    Console.WriteLine("========================\n");
                    
                    // 선택된 범위의 위치 정보
                    string position = $"{selection.Address.Split(',')[0]}";
                    
                    // 선택된 데이터 출력
                    Console.WriteLine("\n=== 선택된 Excel 데이터 ===");
                    Console.WriteLine($"위치: {position}");
                    Console.WriteLine($"내용:\n{htmlContent}");
                    Console.WriteLine("========================\n");

                    // LogWindow의 텍스트박스 업데이트
                    LogWindow.Instance.Dispatcher.Invoke(() =>
                    {
                        LogWindow.Instance.FilePathTextBox.Text = _workbook.FullName;
                        LogWindow.Instance.PositionTextBox.Text = position;
                        LogWindow.Instance.ContextTextBox.Text = htmlContent;
                    });

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

                        string fullHtml = string.Format(htmlTemplate, htmlContent, htmlContent.Length);
                        File.WriteAllText("test.html", fullHtml);
                        Console.WriteLine("test.html 파일이 성공적으로 업데이트되었습니다.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"test.html 파일 업데이트 실패: {ex.Message}");
                    }
                    
                    return (htmlContent, styleAttributes, position);
                }
                else
                {
                    Console.WriteLine("클립보드에 HTML 형식이 없습니다.");
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Excel 데이터 읽기 오류: {ex.Message}");
                Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                return (string.Empty, new Dictionary<string, object>(), string.Empty);
            }
            finally
            {
                // COM 객체 해제
                if (_workbook != null) Marshal.ReleaseComObject(_workbook);
                if (_excelApp != null) Marshal.ReleaseComObject(_excelApp);
            }
        }

        public override (ulong? FileId, uint? VolumeId, string FileType, string FileName, string FilePath) GetFileInfo()
        {
            Microsoft.Office.Interop.Excel.Application? tempExcelApp = null;
            Workbook? tempWorkbook = null;
            
            try
            {
                tempExcelApp = (Microsoft.Office.Interop.Excel.Application)GetActiveObject("Excel.Application");
                tempWorkbook = tempExcelApp.ActiveWorkbook;
                
                if (tempWorkbook == null)
                {
                    return (null, null, "Excel", string.Empty, string.Empty);
                }

                string filePath = tempWorkbook.FullName;
                string fileName = tempWorkbook.Name;
                
                // filePath가 비어있지 않을 때 일치 여부 체크
                if (!string.IsNullOrEmpty(_filePath) && !string.IsNullOrEmpty(filePath))
                {
                    if (!string.Equals(_filePath, filePath, StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine($"파일 경로가 일치하지 않습니다. 기대: {_filePath}, 실제: {filePath}");
                        return (null, null, "Excel", string.Empty, string.Empty);
                    }
                }
                
                // 파일 정보 가져오기
                FileInfo fileInfo = new FileInfo(filePath);
                ulong fileId = (ulong)fileInfo.GetHashCode();
                uint volumeId = (uint)(fileInfo.Directory?.Root.GetHashCode() ?? 0);
                
                return (fileId, volumeId, "Excel", fileName, filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Excel 파일 정보 가져오기 오류: {ex.Message}");
                return (null, null, "Excel", string.Empty, string.Empty);
            }
            finally
            {
                if (tempWorkbook != null) Marshal.ReleaseComObject(tempWorkbook);
                if (tempExcelApp != null) Marshal.ReleaseComObject(tempExcelApp);
            }
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
    }
}