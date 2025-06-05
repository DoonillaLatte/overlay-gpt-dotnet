using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using System.Drawing;                     // ColorTranslator.FromOle() 용
using HtmlAgilityPack;                     // HtmlAgilityPack.HtmlDocument
using ExCSS;                               // StylesheetParser, StyleRule

// 충돌 방지를 위해 별칭(alias) 지정
using ExcelRange = Microsoft.Office.Interop.Excel.Range;
using HapHtmlDocument = HtmlAgilityPack.HtmlDocument;

namespace overlay_gpt
{
    public class ExcelContextReader : BaseContextReader
    {
        private ExcelApp? _excelApp;
        private Workbook? _workbook;
        private Worksheet? _worksheet;

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("oleaut32.dll")]
        private static extern int GetActiveObject(
            ref Guid rclsid,
            IntPtr pvReserved,
            [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        [DllImport("ole32.dll")]
        private static extern int CLSIDFromProgID(
            [MarshalAs(UnmanagedType.LPWStr)] string lpszProgID,
            out Guid pclsid);

        // NTFS 파일 ID를 가져오기 위한 Windows API
        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr CreateFile(
            string lpFileName,
            uint dwDesiredAccess,
            uint dwShareMode,
            IntPtr lpSecurityAttributes,
            uint dwCreationDisposition,
            uint dwFlagsAndAttributes,
            IntPtr hTemplateFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool GetFileInformationByHandle(
            IntPtr hFile,
            out BY_HANDLE_FILE_INFORMATION lpFileInformation);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr hObject);

        [StructLayout(LayoutKind.Sequential)]
        private struct BY_HANDLE_FILE_INFORMATION
        {
            public uint dwFileAttributes;
            public FILETIME ftCreationTime;
            public FILETIME ftLastAccessTime;
            public FILETIME ftLastWriteTime;
            public uint dwVolumeSerialNumber;
            public uint nFileSizeHigh;
            public uint nFileSizeLow;
            public uint nNumberOfLinks;
            public uint nFileIndexHigh;
            public uint nFileIndexLow;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct FILETIME
        {
            public uint dwLowDateTime;
            public uint dwHighDateTime;
        }

        private const uint GENERIC_READ = 0x80000000;
        private const uint FILE_SHARE_READ = 0x00000001;
        private const uint FILE_SHARE_WRITE = 0x00000002;
        private const uint OPEN_EXISTING = 3;

        private static object GetActiveObject(string progID)
        {
            Guid clsid;
            CLSIDFromProgID(progID, out clsid);
            object obj;
            GetActiveObject(ref clsid, IntPtr.Zero, out obj);
            return obj;
        }

        // 파일 ID를 가져오는 메서드
        private (ulong FileId, uint VolumeId)? GetFileId(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                    return null;

                IntPtr handle = CreateFile(
                    filePath,
                    GENERIC_READ,
                    FILE_SHARE_READ | FILE_SHARE_WRITE,
                    IntPtr.Zero,
                    OPEN_EXISTING,
                    0,
                    IntPtr.Zero);

                if (handle.ToInt64() == -1)
                    return null;

                try
                {
                    BY_HANDLE_FILE_INFORMATION fileInfo;
                    if (GetFileInformationByHandle(handle, out fileInfo))
                    {
                        ulong fileId = ((ulong)fileInfo.nFileIndexHigh << 32) | fileInfo.nFileIndexLow;
                        return (fileId, fileInfo.dwVolumeSerialNumber);
                    }
                }
                finally
                {
                    CloseHandle(handle);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"파일 ID 가져오기 오류: {ex.Message}");
            }
            return null;
        }

        private string GetBorderStyleString(Dictionary<string, object> borderInfo)
        {
            if (!borderInfo.ContainsKey("LineStyle") 
                || borderInfo["LineStyle"] == null 
                || borderInfo["LineStyle"] == DBNull.Value)
            {
                return string.Empty;
            }

            int lineStyle = SafeGetInt(borderInfo["LineStyle"]);
            int weight    = SafeGetInt(borderInfo["Weight"]);
            int colorOle  = SafeGetInt(borderInfo["Color"]);

            string style = lineStyle switch
            {
                -4142 => "none",      // 없음
                1     => "1px solid", // 실선
                2     => "2px solid",
                4     => "1px dashed",
                5     => "1px dotted",
                _     => "none"
            };

            if (style == "none")
                return string.Empty;

            // OLE 색을 System.Drawing.Color로 변환
            System.Drawing.Color col = ColorTranslator.FromOle(colorOle);
            string colorHex = $"#{col.R:X2}{col.G:X2}{col.B:X2}";

            return $"{style} {colorHex}";
        }

        private bool SafeGetBoolean(object value)
        {
            if (value == null || value == DBNull.Value)
                return false;
            return Convert.ToBoolean(value);
        }

        private int SafeGetInt(object value, int defaultValue = 0)
        {
            if (value == null || value == DBNull.Value)
                return defaultValue;
            return Convert.ToInt32(value);
        }

        // 셀 텍스트 꾸미기(볼드, 이탤릭 등만 적용)
        private string GetStyledText(string text, Dictionary<string, object> styleAttributes)
        {
            string result = text;
            if (styleAttributes.ContainsKey("UnderlineStyle") 
                && styleAttributes["UnderlineStyle"]?.ToString() == "Single")
            {
                result = $"<u>{result}</u>";
            }
            if (styleAttributes.ContainsKey("FontWeight") 
                && styleAttributes["FontWeight"]?.ToString() == "Bold")
            {
                result = $"<b>{result}</b>";
            }
            if (styleAttributes.ContainsKey("FontItalic") 
                && SafeGetBoolean(styleAttributes["FontItalic"]))
            {
                result = $"<i>{result}</i>";
            }
            if (styleAttributes.ContainsKey("FontStrikethrough") 
                && SafeGetBoolean(styleAttributes["FontStrikethrough"]))
            {
                result = $"<s>{result}</s>";
            }
            if (styleAttributes.ContainsKey("FontSuperscript") 
                && SafeGetBoolean(styleAttributes["FontSuperscript"]))
            {
                result = $"<sup>{result}</sup>";
            }
            if (styleAttributes.ContainsKey("FontSubscript") 
                && SafeGetBoolean(styleAttributes["FontSubscript"]))
            {
                result = $"<sub>{result}</sub>";
            }
            return result;
        }

        // 셀 스타일 문자열 추출 (background-color, color 포함)
        private string GetCellStyleString(Dictionary<string, object> styleAttributes)
        {
            var styleList = new List<string>();

            // ─── 배경색 처리 ───
            if (styleAttributes.ContainsKey("BackgroundColor"))
            {
                int oleColor = Convert.ToInt32(styleAttributes["BackgroundColor"]);
                if (oleColor != 16777215) // 기본 흰색이 아닌 경우에만
                {
                    System.Drawing.Color bgCol = ColorTranslator.FromOle(oleColor);
                    string hexColor = $"#{bgCol.R:X2}{bgCol.G:X2}{bgCol.B:X2}";
                    styleList.Add($"background-color: {hexColor}");
                }
            }

            // ─── 글자색 처리 ───
            if (styleAttributes.ContainsKey("ForegroundColor"))
            {
                int oleColor = Convert.ToInt32(styleAttributes["ForegroundColor"]);
                if (oleColor != 0) // 기본 검정색이 아닌 경우에만
                {
                    System.Drawing.Color fgCol = ColorTranslator.FromOle(oleColor);
                    string hexColor = $"#{fgCol.R:X2}{fgCol.G:X2}{fgCol.B:X2}";
                    styleList.Add($"color: {hexColor}");
                }
            }

            // ─── 폰트 처리 ───
            if (styleAttributes.ContainsKey("FontName"))
            {
                string fontName = styleAttributes["FontName"]?.ToString() ?? string.Empty;
                if (!string.IsNullOrEmpty(fontName))
                {
                    styleList.Add($"font-family: {fontName}");
                }
            }

            // ─── 폰트 크기 처리 ───
            if (styleAttributes.ContainsKey("FontSize"))
            {
                double fontSize = Convert.ToDouble(styleAttributes["FontSize"]);
                styleList.Add($"font-size: {fontSize}pt");
            }

            // ─── 가로 정렬 처리 ───
            if (styleAttributes.ContainsKey("HorizontalAlignment"))
            {
                string alignment = styleAttributes["HorizontalAlignment"]?.ToString() ?? "Left";
                string cssAlignment = alignment switch
                {
                    "Center"  => "center",
                    "Right"   => "right",
                    "Justify" => "justify",
                    _         => "left"
                };
                styleList.Add($"text-align: {cssAlignment}");
            }

            // ─── 세로 정렬 처리 ───
            if (styleAttributes.ContainsKey("VerticalAlignment"))
            {
                string alignment = styleAttributes["VerticalAlignment"]?.ToString() ?? "Bottom";
                string cssAlignment = alignment switch
                {
                    "Top"    => "top",
                    "Center" => "middle",
                    _        => "bottom"
                };
                styleList.Add($"vertical-align: {cssAlignment}");
            }

            // ─── 들여쓰기 처리 ───
            if (styleAttributes.ContainsKey("IndentLevel"))
            {
                int indentLevel = Convert.ToInt32(styleAttributes["IndentLevel"]);
                if (indentLevel > 0)
                {
                    styleList.Add($"padding-left: {indentLevel * 20}px");
                }
            }

            // ─── 테두리 처리 ───
            if (styleAttributes.ContainsKey("BorderStyle"))
            {
                var borders = (Borders)styleAttributes["BorderStyle"];
                var borderPositions = new Dictionary<string, XlBordersIndex>
                {
                    { "top",    XlBordersIndex.xlEdgeTop },
                    { "right",  XlBordersIndex.xlEdgeRight },
                    { "bottom", XlBordersIndex.xlEdgeBottom },
                    { "left",   XlBordersIndex.xlEdgeLeft }
                };

                foreach (var position in borderPositions)
                {
                    var border = borders[position.Value];
                    if (border != null)
                    {
                        var borderInfo = new Dictionary<string, object>
                        {
                            { "LineStyle", border.LineStyle },
                            { "Weight",    border.Weight },
                            { "Color",     border.Color }
                        };
                        string borderStyle = GetBorderStyleString(borderInfo);
                        if (!string.IsNullOrEmpty(borderStyle))
                        {
                            styleList.Add($"border-{position.Key}: {borderStyle}");
                        }
                    }
                }
            }

            // ─── 셀 크기 처리 ───
            if (styleAttributes.ContainsKey("Width"))
            {
                double width = Convert.ToDouble(styleAttributes["Width"]);
                styleList.Add($"width: {width}pt");
            }

            if (styleAttributes.ContainsKey("Height"))
            {
                double height = Convert.ToDouble(styleAttributes["Height"]);
                styleList.Add($"height: {height}pt");
            }

            return string.Join("; ", styleList);
        }

        public override (string SelectedText, Dictionary<string, object> StyleAttributes, string LineNumber) GetSelectedTextWithStyle(bool readAllContent = false)
        {
            try
            {
                Console.WriteLine("Excel 데이터 읽기 시작...");

                var excelProcesses = Process.GetProcessesByName("EXCEL");
                if (excelProcesses.Length == 0)
                {
                    Console.WriteLine("실행 중인 Excel 애플리케이션을 찾을 수 없습니다.");
                    throw new InvalidOperationException("Excel is not running");
                }

                Process? activeExcelProcess = null;
                foreach (var process in excelProcesses)
                {
                    if (process.MainWindowHandle != IntPtr.Zero 
                        && process.MainWindowTitle.Length > 0)
                    {
                        Console.WriteLine($"Excel 프로세스 정보:");
                        Console.WriteLine($"- 프로세스 ID: {process.Id}");
                        Console.WriteLine($"- 창 제목: {process.MainWindowTitle}");
                        Console.WriteLine($"- 실행 경로: {process.MainModule?.FileName}");

                        if (process.MainWindowHandle == GetForegroundWindow())
                        {
                            activeExcelProcess = process;
                            Console.WriteLine("이 Excel 창이 현재 활성화되어 있습니다.");
                        }
                    }
                }

                if (activeExcelProcess == null)
                {
                    Console.WriteLine("활성화된 Excel 창을 찾을 수 없습니다.");
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);
                }

                try
                {
                    Console.WriteLine("Excel COM 객체 가져오기 시도...");
                    _excelApp = (ExcelApp)GetActiveObject("Excel.Application");
                    Console.WriteLine("Excel COM 객체 가져오기 성공");

                    Console.WriteLine("활성 워크북 가져오기 시도...");
                    _workbook = _excelApp.ActiveWorkbook;

                    if (_workbook == null)
                    {
                        Console.WriteLine("활성 워크북을 찾을 수 없습니다.");
                        return (string.Empty, new Dictionary<string, object>(), string.Empty);
                    }

                    Console.WriteLine($"활성 워크북 정보:");
                    Console.WriteLine($"- 워크북 이름: {_workbook.Name}");
                    Console.WriteLine($"- 전체 경로: {_workbook.FullName}");
                    Console.WriteLine($"- 저장 여부: {(_workbook.Saved ? "저장됨" : "저장되지 않음")}");
                    Console.WriteLine($"- 읽기 전용: {(_workbook.ReadOnly ? "예" : "아니오")}");

                    _worksheet = _excelApp.ActiveSheet as Worksheet;
                    if (_worksheet == null)
                    {
                        Console.WriteLine("활성 워크시트를 찾을 수 없습니다.");
                        return (string.Empty, new Dictionary<string, object>(), string.Empty);
                    }

                    var range = readAllContent
                        ? _worksheet.UsedRange
                        : _excelApp.Selection as ExcelRange;
                    if (range == null)
                    {
                        Console.WriteLine("선택된 셀이 없습니다.");
                        return (string.Empty, new Dictionary<string, object>(), string.Empty);
                    }

                    // 선택된 범위의 시작 행/열과 끝 행/열 구하기
                    int startRow = range.Row;
                    int endRow   = range.Row + range.Rows.Count - 1;
                    int startCol = range.Column;
                    int endCol   = range.Column + range.Columns.Count - 1;
                    string lineNumber = $"R{startRow}C{startCol}-R{endRow}C{endCol}";

                    // HTML 형식으로 클립보드 복사 시도
                    try
                    {
                        range.Copy();
                        if (Clipboard.ContainsText(TextDataFormat.Html))
                        {
                            string htmlContent = Clipboard.GetText(TextDataFormat.Html);

                            // 1) <!--StartFragment--> 와 <!--EndFragment--> 사이의 순수 HTML만 추출
                            int startIdx = htmlContent.IndexOf("<!--StartFragment-->");
                            int endIdx   = htmlContent.IndexOf("<!--EndFragment-->");
                            if (startIdx != -1 && endIdx != -1 && endIdx > startIdx)
                            {
                                int fragContentStart = startIdx + "<!--StartFragment-->".Length;
                                int fragLength = endIdx - fragContentStart;
                                string rawFragment = htmlContent.Substring(fragContentStart, fragLength);

                                // 1-1) style 블록 전체 추출
                                string styleBlock = "";
                                {
                                    var fullDoc = new HapHtmlDocument();
                                    fullDoc.LoadHtml(htmlContent);
                                    var styleNode = fullDoc.DocumentNode.SelectSingleNode("//style");
                                    if (styleNode != null)
                                    {
                                        styleBlock = styleNode.OuterHtml;
                                    }
                                }

                                // 2) 불필요 태그/속성 제거 및 인접 노드 병합
                                //    (클래스 → 인라인 스타일 처리 포함)
                                string cleanedHtml = CleanAndNormalizeHtml(rawFragment, styleBlock);

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

                                    string fullHtml = string.Format(htmlTemplate, cleanedHtml, cleanedHtml.Length);
                                    File.WriteAllText("test.html", fullHtml);
                                    Console.WriteLine("test.html 파일이 성공적으로 업데이트되었습니다.");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"test.html 파일 업데이트 실패: {ex.Message}");
                                }

                                // 로그 윈도우의 컨텍스트 텍스트박스 업데이트
                                LogWindow.Instance.Dispatcher.Invoke(() =>
                                {
                                    LogWindow.Instance.FilePathTextBox.Text = _workbook.FullName;
                                    LogWindow.Instance.PositionTextBox.Text = lineNumber;
                                    LogWindow.Instance.ContextTextBox.Text = cleanedHtml;
                                });

                                return (cleanedHtml, new Dictionary<string, object>(), lineNumber);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"HTML 클립보드 복사 실패: {ex.Message}");
                    }

                    // HTML 형식으로 가져오기 실패 시 기존 방식으로 처리
                    var tableHtml = new StringBuilder();
                    tableHtml.Append("<table style='border-collapse: collapse;'>");

                    int currentRow = -1;
                    foreach (ExcelRange cell in range)
                    {
                        if (cell.Row != currentRow)
                        {
                            if (currentRow != -1)
                            {
                                tableHtml.Append("</tr>");
                            }
                            tableHtml.Append("<tr>");
                            currentRow = cell.Row;
                        }

                        string cellText = cell.Text?.ToString() ?? string.Empty;
                        var cellStyle = new Dictionary<string, object>();

                        try
                        {
                            cellStyle["FontName"]        = cell.Font?.Name ?? "Calibri";
                            cellStyle["FontSize"]        = cell.Font?.Size ?? 11;
                            cellStyle["FontWeight"]      = (cell.Font?.Bold ?? false) ? "Bold" : "Normal";
                            cellStyle["FontItalic"]      = cell.Font?.Italic ?? false;
                            cellStyle["ForegroundColor"] = cell.Font?.Color ?? 0;
                            cellStyle["BackgroundColor"] = cell.Interior?.Color ?? 16777215;
                            cellStyle["BorderStyle"]     = cell.Borders;
                            cellStyle["Width"]           = cell.Width;
                            cellStyle["Height"]          = cell.Height;

                            string styleString = GetCellStyleString(cellStyle);
                            cellText = GetStyledText(cellText, cellStyle);

                            tableHtml.Append($"<td style='{styleString}'>{cellText}</td>");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"셀 스타일 변환 중 오류 발생: {ex.Message}");
                            tableHtml.Append($"<td>{cellText}</td>");
                        }
                    }

                    if (currentRow != -1)
                    {
                        tableHtml.Append("</tr>");
                    }
                    tableHtml.Append("</table>");

                    string selectedText = tableHtml.ToString();

                    // 로그 윈도우의 컨텍스트 텍스트박스 업데이트
                    LogWindow.Instance.Dispatcher.Invoke(() =>
                    {
                        LogWindow.Instance.FilePathTextBox.Text = _workbook.FullName;
                        LogWindow.Instance.PositionTextBox.Text = lineNumber;
                        LogWindow.Instance.ContextTextBox.Text = selectedText;
                    });

                    return (selectedText, new Dictionary<string, object>(), lineNumber);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Excel COM 연결 오류: {ex.Message}");
                    return (string.Empty, new Dictionary<string, object>(), string.Empty);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Excel 데이터 읽기 오류 발생: {ex.Message}");
                Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                LogWindow.Instance.Log($"Excel 데이터 읽기 오류: {ex.Message}");
                return (string.Empty, new Dictionary<string, object>(), string.Empty);
            }
            finally
            {
                if (_worksheet != null) Marshal.ReleaseComObject(_worksheet);
                if (_workbook  != null) Marshal.ReleaseComObject(_workbook);
                if (_excelApp  != null) Marshal.ReleaseComObject(_excelApp);
            }
        }

        private string CleanAndNormalizeHtml(string rawFragment, string styleBlock)
        {
            // 0) styleBlock에서 순수 CSS 텍스트만 추출
            string cssContent = "";
            if (!string.IsNullOrEmpty(styleBlock))
            {
                int start = styleBlock.IndexOf('>');
                int end   = styleBlock.LastIndexOf("</style>", StringComparison.OrdinalIgnoreCase);
                if (start >= 0 && end > start)
                {
                    cssContent = styleBlock.Substring(start + 1, end - (start + 1));
                }
            }

            // 1) CSS 파싱: ExCSS를 이용해서 "클래스명 → 인라인 스타일 문자열" 딕셔너리 생성
            var parser = new StylesheetParser();
            var stylesheet = parser.Parse(cssContent);

            var classToStyleDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // StyleSheet의 Rules를 사용하여 클래스 → 인라인 스타일 문자열 생성
            foreach (StyleRule rule in stylesheet.StyleRules)
            {
                string selector = rule.Selector.ToString().Trim();
                if (selector.StartsWith("."))
                {
                    string className = selector.Substring(1);
                    var sbDecl = new StringBuilder();

                    foreach (var decl in rule.Style)
                    {
                        sbDecl.Append($"{decl.Name}: {decl.Value}; ");
                    }

                    classToStyleDict[className] = sbDecl.ToString().Trim();
                }
            }

            // 2) HtmlAgilityPack으로 rawFragment를 로드
            var htmlDoc = new HapHtmlDocument();
            htmlDoc.LoadHtml("<div id=\"wrapper\">" + rawFragment + "</div>");
            var wrapper = htmlDoc.GetElementbyId("wrapper")!;

            // 3) 불필요 노드 제거
            RemoveUnwantedNodes(wrapper);
            RemoveExcelSpecificAttributes(wrapper);

            // 4) 클래스→인라인 스타일 치환
            var nodesWithClass = wrapper.SelectNodes("//*[@class]");
            if (nodesWithClass != null)
            {
                foreach (var node in nodesWithClass.ToList())
                {
                    string? classAttr = node.GetAttributeValue("class", null);
                    if (string.IsNullOrEmpty(classAttr))
                        continue;

                    var classNames = classAttr
                        .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    var existingStyle = node.GetAttributeValue("style", "").Trim();

                    var sbStyle = new StringBuilder(existingStyle);
                    if (existingStyle.Length > 0 && !existingStyle.EndsWith(";"))
                        sbStyle.Append("; ");

                    foreach (var cls in classNames)
                    {
                        if (classToStyleDict.TryGetValue(cls, out string styleFromClass))
                        {
                            sbStyle.Append(styleFromClass);
                            if (!styleFromClass.EndsWith(";"))
                                sbStyle.Append("; ");
                        }
                    }

                    var finalStyle = sbStyle.ToString().Trim();
                    if (!string.IsNullOrEmpty(finalStyle))
                    {
                        node.SetAttributeValue("style", finalStyle);
                    }

                    node.Attributes.Remove("class");
                }
            }

            // 5) 이후 기존 로직 이어감
            RemoveEmptySpans(wrapper);
            MergeAdjacentSpans(wrapper);
            RemoveEmptyStyleAttributes(wrapper);
            PreserveTableStyles(wrapper);

            // 6) 최종 HTML 문자열 추출 및 정리
            string interimHtml = wrapper.InnerHtml;
            string normalized = NormalizeWhitespace(interimHtml);
            return normalized;
        }

        private void RemoveUnwantedNodes(HtmlNode root)
        {
            // 1) <meta> 태그 전부 제거
            var metas = root.SelectNodes("//meta");
            if (metas != null)
            {
                foreach (var meta in metas)
                    meta.Remove();
            }

            // 2) <xml> ... </xml> 노드 (Office 전용) 제거
            var xmlNodes = root.SelectNodes("//xml");
            if (xmlNodes != null)
            {
                foreach (var node in xmlNodes)
                    node.Remove();
            }

            // 3) Excel 고유 네임스페이스가 들어간 모든 노드 제거
            var allNodes = root.SelectNodes("//*");
            if (allNodes != null)
            {
                foreach (var node in allNodes.ToList())
                {
                    if (node.Name.StartsWith("o:", StringComparison.OrdinalIgnoreCase)
                        || node.Name.StartsWith("x:", StringComparison.OrdinalIgnoreCase)
                        || node.Name.StartsWith("v:", StringComparison.OrdinalIgnoreCase))
                    {
                        node.Remove();
                    }
                }
            }

            // 4) 조건부 주석 제거
            var comments = root.SelectNodes("//comment()");
            if (comments != null)
            {
                foreach (var commentNode in comments.Cast<HtmlCommentNode>())
                {
                    commentNode.Remove();
                }
            }
        }

        private void RemoveExcelSpecificAttributes(HtmlNode root)
        {
            // 1) 모든 노드를 순회하면서 lang 속성을 제거
            var nodesWithLang = root.SelectNodes("//*[@lang]");
            if (nodesWithLang != null)
            {
                foreach (var node in nodesWithLang)
                {
                    node.Attributes.Remove("lang");
                }
            }

            // 2) style 속성이 있는 노드에서 Excel 전용 속성들을 삭제
            var nodesWithStyle = root.SelectNodes("//*[@style]");
            if (nodesWithStyle != null)
            {
                foreach (var node in nodesWithStyle.ToList())
                {
                    var styleAttr = node.GetAttributeValue("style", "").Trim();
                    if (string.IsNullOrEmpty(styleAttr))
                    {
                        node.Attributes.Remove("style");
                        continue;
                    }

                    var declarations = styleAttr
                        .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(decl => decl.Trim())
                        .Where(decl => !decl.StartsWith("mso-", StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    if (declarations.Any())
                    {
                        string newStyleValue = string.Join(";", declarations) + ";";
                        node.SetAttributeValue("style", newStyleValue);
                    }
                    else
                    {
                        node.Attributes.Remove("style");
                    }
                }
            }
        }

        private void RemoveEmptySpans(HtmlNode root)
        {
            var spans = root.SelectNodes("//span");
            if (spans != null)
            {
                foreach (var span in spans.ToList())
                {
                    string styleAttr = span.GetAttributeValue("style", "").Trim();
                    string inner = span.InnerHtml.Trim();
                    if (string.IsNullOrEmpty(styleAttr) && string.IsNullOrEmpty(inner))
                    {
                        span.Remove();
                    }
                }
            }
        }

        private void MergeAdjacentSpans(HtmlNode root)
        {
            var parentNodes = root.SelectNodes("//*");
            if (parentNodes == null) return;

            foreach (var parent in parentNodes)
            {
                var children = parent.ChildNodes.ToList();
                for (int i = 0; i < children.Count - 1; i++)
                {
                    var curr = children[i];
                    var next = children[i + 1];

                    if (curr.Name.Equals("span", StringComparison.OrdinalIgnoreCase)
                        && next.Name.Equals("span", StringComparison.OrdinalIgnoreCase))
                    {
                        string styleCurr = curr.GetAttributeValue("style", "");
                        string styleNext = next.GetAttributeValue("style", "");
                        if (styleCurr == styleNext)
                        {
                            curr.InnerHtml = curr.InnerHtml + next.InnerHtml;
                            next.Remove();
                            children = parent.ChildNodes.ToList();
                            i--;
                        }
                    }
                }
            }
        }

        private void RemoveEmptyStyleAttributes(HtmlNode root)
        {
            var nodesWithStyle = root.SelectNodes("//*[@style]");
            if (nodesWithStyle != null)
            {
                foreach (var node in nodesWithStyle.ToList())
                {
                    var val = node.GetAttributeValue("style", "").Trim();
                    if (string.IsNullOrEmpty(val))
                        node.Attributes.Remove("style");
                }
            }
        }

        private void PreserveTableStyles(HtmlNode root)
        {
            // 테이블 스타일 보존
            var tables = root.SelectNodes("//table");
            if (tables != null)
            {
                foreach (var table in tables)
                {
                    table.SetAttributeValue("style", "border-collapse: collapse; border: 1px solid #000000;");
                }
            }

            // 행 스타일 보존
            var rows = root.SelectNodes("//tr");
            if (rows != null)
            {
                foreach (var row in rows)
                {
                    string height = row.GetAttributeValue("height", "");
                    if (!string.IsNullOrEmpty(height))
                    {
                        row.SetAttributeValue("style", $"height: {height};");
                    }
                }
            }

            // 셀 스타일 보존
            var cells = root.SelectNodes("//td");
            if (cells != null)
            {
                foreach (var cell in cells)
                {
                    var styleList = new List<string>();

                    // 너비
                    string width = cell.GetAttributeValue("width", "");
                    if (!string.IsNullOrEmpty(width))
                    {
                        styleList.Add($"width: {width}");
                    }

                    // 높이
                    string height = cell.GetAttributeValue("height", "");
                    if (!string.IsNullOrEmpty(height))
                    {
                        styleList.Add($"height: {height}");
                    }

                    // 기존 스타일
                    string existingStyle = cell.GetAttributeValue("style", "");
                    if (!string.IsNullOrEmpty(existingStyle))
                    {
                        styleList.Add(existingStyle);
                    }

                    // 테두리 기본값 추가
                    styleList.Add("border: 1px solid #000000");

                    if (styleList.Count > 0)
                    {
                        cell.SetAttributeValue("style", string.Join("; ", styleList));
                    }
                }
            }
        }

        private string NormalizeWhitespace(string html)
        {
            if (string.IsNullOrEmpty(html))
                return string.Empty;

            return html.Trim();
        }

        public override (ulong? FileId, uint? VolumeId, string FileType, string FileName, string FilePath) GetFileInfo()
        {
            ExcelApp? tempExcelApp = null;
            Workbook? tempWorkbook = null;

            try
            {
                Console.WriteLine("Excel COM 객체 가져오기 시도...");
                tempExcelApp = (ExcelApp)GetActiveObject("Excel.Application");
                Console.WriteLine("Excel COM 객체 가져오기 성공");

                Console.WriteLine("활성 워크북 가져오기 시도...");
                tempWorkbook = tempExcelApp.ActiveWorkbook;

                if (tempWorkbook == null)
                {
                    Console.WriteLine("활성 워크북을 찾을 수 없습니다.");
                    return (null, null, "Excel", string.Empty, string.Empty);
                }

                string filePath = tempWorkbook.FullName;
                string fileName = tempWorkbook.Name;

                Console.WriteLine($"Excel 문서 정보:");
                Console.WriteLine($"- 파일 경로: {filePath}");
                Console.WriteLine($"- 파일 이름: {fileName}");

                if (string.IsNullOrEmpty(filePath))
                {
                    Console.WriteLine("파일 경로가 비어있습니다.");
                    return (null, null, "Excel", fileName, string.Empty);
                }

                var fileIdInfo = GetFileId(filePath);

                if (fileIdInfo == null)
                {
                    Console.WriteLine("파일 ID 정보를 가져오지 못했습니다.");
                }
                else
                {
                    Console.WriteLine($"파일 ID 정보:");
                    Console.WriteLine($"- FileId: {fileIdInfo.Value.FileId}");
                    Console.WriteLine($"- VolumeId: {fileIdInfo.Value.VolumeId}");
                }

                return (
                    fileIdInfo?.FileId,
                    fileIdInfo?.VolumeId,
                    "Excel",
                    fileName,
                    filePath
                );
            }
            catch (Exception ex)
            {
                Console.WriteLine($"파일 정보 가져오기 오류: {ex.Message}");
                Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                return (null, null, "Excel", string.Empty, string.Empty);
            }
            finally
            {
                if (tempWorkbook != null) Marshal.ReleaseComObject(tempWorkbook);
                if (tempExcelApp != null) Marshal.ReleaseComObject(tempExcelApp);
            }
        }
    }
}
