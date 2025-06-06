using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using System.IO;

namespace overlay_gpt
{
    public class ExcelContextReader : BaseContextReader
    {
        private Microsoft.Office.Interop.Excel.Application? _excelApp;
        private Workbook? _workbook;

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

        public override (string SelectedText, Dictionary<string, object> StyleAttributes, string LineNumber) GetSelectedTextWithStyle(bool readAllContent = false)
        {
            try
            {
                Console.WriteLine("Excel 데이터 읽기 시작...");

                // Excel COM 객체 가져오기
                _excelApp = (Microsoft.Office.Interop.Excel.Application)GetActiveObject("Excel.Application");
                if (_excelApp == null)
                {
                    Console.WriteLine("Excel 애플리케이션을 찾을 수 없습니다.");
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);
                }

                // 활성 워크북 가져오기
                _workbook = _excelApp.ActiveWorkbook;
                if (_workbook == null)
                {
                    Console.WriteLine("활성 워크북을 찾을 수 없습니다.");
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);
                }

                // 선택된 범위 가져오기
                Microsoft.Office.Interop.Excel.Range? selection = _excelApp.Selection as Microsoft.Office.Interop.Excel.Range;
                if (selection == null)
                {
                    Console.WriteLine("선택된 범위가 없습니다.");
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);
                }

                // 선택된 범위를 클립보드에 복사
                selection.Copy();

                // HTML 형식으로 클립보드에서 가져오기
                if (Clipboard.ContainsText(TextDataFormat.Html))
                {
                    string htmlContent = Clipboard.GetText(TextDataFormat.Html);
                    
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
                
                return (null, null, "Excel", fileName, filePath);
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