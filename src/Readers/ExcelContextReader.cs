using System;
using System.Collections.Generic;
using System.Windows.Automation;
using System.Windows.Automation.Text;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.IO;

namespace overlay_gpt
{
    public class ExcelContextReader : BaseContextReader
    {
        private Application? _excelApp;
        private Workbook? _workbook;
        private Worksheet? _worksheet;

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("oleaut32.dll")]
        private static extern int GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        [DllImport("ole32.dll")]
        private static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);

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
            if (!borderInfo.ContainsKey("LineStyle") || borderInfo["LineStyle"] == null || borderInfo["LineStyle"] == DBNull.Value)
                return string.Empty;

            int lineStyle = SafeGetInt(borderInfo["LineStyle"]);
            int color = SafeGetInt(borderInfo["Color"]);

            string style = lineStyle switch
            {
                -4142 => "none",      // 없음
                1     => "1px solid", // 실선
                2     => "2px solid",
                4     => "1px dashed",
                5     => "1px dotted",
                // 필요시 추가
                _     => "none"
            };

            if (style == "none")
                return string.Empty;

            string colorHex = $"#{color:X6}";
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

        // BGR 색상값을 RGB로 변환
        private int ConvertBGRToRGB(int bgrColor)
        {
            int r = (bgrColor >> 16) & 0xFF;
            int g = (bgrColor >> 8) & 0xFF;
            int b = bgrColor & 0xFF;
            return (b << 16) | (g << 8) | r;
        }

        // 셀 텍스트 꾸미기(볼드, 이탤릭 등만 적용)
        private string GetStyledText(string text, Dictionary<string, object> styleAttributes)
        {
            string result = text;
            if (styleAttributes.ContainsKey("UnderlineStyle") && styleAttributes["UnderlineStyle"]?.ToString() == "Single")
                result = $"<u>{result}</u>";
            if (styleAttributes.ContainsKey("FontWeight") && styleAttributes["FontWeight"]?.ToString() == "Bold")
                result = $"<b>{result}</b>";
            if (styleAttributes.ContainsKey("FontItalic") && SafeGetBoolean(styleAttributes["FontItalic"]))
                result = $"<i>{result}</i>";
            if (styleAttributes.ContainsKey("FontStrikethrough") && SafeGetBoolean(styleAttributes["FontStrikethrough"]))
                result = $"<s>{result}</s>";
            if (styleAttributes.ContainsKey("FontSuperscript") && SafeGetBoolean(styleAttributes["FontSuperscript"]))
                result = $"<sup>{result}</sup>";
            if (styleAttributes.ContainsKey("FontSubscript") && SafeGetBoolean(styleAttributes["FontSubscript"]))
                result = $"<sub>{result}</sub>";
            return result;
        }

        // 셀 스타일 문자열 추출 (border 포함)
        private string GetCellStyleString(Dictionary<string, object> styleAttributes)
        {
            var styleList = new List<string>();
            if (styleAttributes.ContainsKey("BackgroundColor"))
            {
                int bgColor = Convert.ToInt32(styleAttributes["BackgroundColor"]);
                if (bgColor != 16777215)
                {
                    string hexColor = $"#{ConvertBGRToRGB(bgColor):X6}";
                    styleList.Add($"background-color: {hexColor}");
                }
            }
            if (styleAttributes.ContainsKey("ForegroundColor"))
            {
                int fgColor = Convert.ToInt32(styleAttributes["ForegroundColor"]);
                if (fgColor != 0)
                {
                    string hexColor = $"#{ConvertBGRToRGB(fgColor):X6}";
                    styleList.Add($"color: {hexColor}");
                }
            }
            if (styleAttributes.ContainsKey("FontName"))
            {
                string fontName = styleAttributes["FontName"]?.ToString() ?? string.Empty;
                if (!string.IsNullOrEmpty(fontName) && fontName != "Calibri" && fontName != "Arial" && fontName != "맑은 고딕")
                {
                    styleList.Add($"font-family: {fontName}");
                }
            }
            if (styleAttributes.ContainsKey("FontSize"))
            {
                double fontSize = Convert.ToDouble(styleAttributes["FontSize"]);
                if (fontSize != 11)
                {
                    styleList.Add($"font-size: {fontSize}pt");
                }
            }
            if (styleAttributes.ContainsKey("HorizontalAlignment"))
            {
                string alignment = styleAttributes["HorizontalAlignment"]?.ToString() ?? "Left";
                if (alignment != "Left")
                {
                    string cssAlignment = alignment switch
                    {
                        "Center" => "center",
                        "Right" => "right",
                        "Justify" => "justify",
                        _ => "left"
                    };
                    styleList.Add($"text-align: {cssAlignment}");
                }
            }
            if (styleAttributes.ContainsKey("VerticalAlignment"))
            {
                string alignment = styleAttributes["VerticalAlignment"]?.ToString() ?? "Bottom";
                if (alignment != "Bottom")
                {
                    string cssAlignment = alignment switch
                    {
                        "Top" => "top",
                        "Center" => "middle",
                        _ => "bottom"
                    };
                    styleList.Add($"vertical-align: {cssAlignment}");
                }
            }
            if (styleAttributes.ContainsKey("IndentLevel"))
            {
                int indentLevel = Convert.ToInt32(styleAttributes["IndentLevel"]);
                if (indentLevel > 0)
                {
                    styleList.Add($"padding-left: {indentLevel * 20}px");
                }
            }
            if (styleAttributes.ContainsKey("BorderStyle"))
            {
                var borderStyles = (Dictionary<string, object>)styleAttributes["BorderStyle"];
                foreach (var border in borderStyles)
                {
                    if (border.Value != null)
                    {
                        var borderInfo = (Dictionary<string, object>)border.Value;
                        string borderStyle = GetBorderStyleString(borderInfo);
                        if (!string.IsNullOrEmpty(borderStyle))
                        {
                            styleList.Add($"border-{border.Key}: {borderStyle}");
                        }
                    }
                }
            }
            if (styleAttributes.ContainsKey("FontShadow") && SafeGetBoolean(styleAttributes["FontShadow"]))
            {
                styleList.Add("text-shadow: 1px 1px 1px rgba(0,0,0,0.5)");
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
                    if (process.MainWindowHandle != IntPtr.Zero && process.MainWindowTitle.Length > 0)
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
                    _excelApp = (Application)GetActiveObject("Excel.Application");
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

                    var range = readAllContent ? _worksheet.UsedRange : _excelApp.Selection as Microsoft.Office.Interop.Excel.Range;
                    if (range == null)
                    {
                        Console.WriteLine("선택된 셀이 없습니다.");
                        return (string.Empty, new Dictionary<string, object>(), string.Empty);
                    }

                    // 선택된 범위의 모든 셀의 텍스트를 수집
                    var tableHtml = new StringBuilder();
                    tableHtml.Append("<table style='border-collapse: collapse;'>");
                    
                    int currentRow = -1;
                    foreach (Microsoft.Office.Interop.Excel.Range cell in range)
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
                            cellStyle["FontName"] = cell.Font?.Name ?? "Calibri";
                            cellStyle["FontSize"] = cell.Font?.Size ?? 11;
                            cellStyle["FontWeight"] = (cell.Font?.Bold ?? false) ? "Bold" : "Normal";
                            cellStyle["FontItalic"] = cell.Font?.Italic ?? false;
                            cellStyle["ForegroundColor"] = cell.Font?.Color ?? 0;
                            cellStyle["BackgroundColor"] = cell.Interior?.Color ?? 16777215;
                            
                            // 셀 스타일 문자열 생성
                            string styleString = GetCellStyleString(cellStyle);
                            
                            // HTML 형식으로 변환
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

                    var styleAttributes = new Dictionary<string, object>();
                    
                    // 스타일 정보 수집
                    try
                    {
                        styleAttributes["FontName"] = range.Font?.Name ?? "Calibri";
                        styleAttributes["FontSize"] = range.Font?.Size ?? 11;
                        styleAttributes["FontWeight"] = (range.Font?.Bold ?? false) ? "Bold" : "Normal";
                        styleAttributes["FontItalic"] = range.Font?.Italic ?? false;
                        styleAttributes["ForegroundColor"] = range.Font?.Color ?? 0;
                        styleAttributes["BackgroundColor"] = range.Interior?.Color ?? 16777215;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"스타일 정보 수집 중 오류 발생: {ex.Message}");
                        // 기본 스타일 정보 설정
                        styleAttributes["FontName"] = "Calibri";
                        styleAttributes["FontSize"] = 11;
                        styleAttributes["FontWeight"] = "Normal";
                        styleAttributes["FontItalic"] = false;
                        styleAttributes["ForegroundColor"] = 0;
                        styleAttributes["BackgroundColor"] = 16777215;
                    }

                    // 선택된 범위의 시작 행/열과 끝 행/열 구하기
                    int startRow = range.Row;
                    int endRow = range.Row + range.Rows.Count - 1;
                    int startCol = range.Column;
                    int endCol = range.Column + range.Columns.Count - 1;
                    string lineNumber = $"R{startRow}C{startCol}-R{endRow}C{endCol}";

                    return (selectedText, styleAttributes, lineNumber);
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
                if (_workbook != null) Marshal.ReleaseComObject(_workbook);
                if (_excelApp != null) Marshal.ReleaseComObject(_excelApp);
            }
        }

        public override (ulong? FileId, uint? VolumeId, string FileType, string FileName, string FilePath) GetFileInfo()
        {
            try
            {
                if (_workbook == null)
                    return (null, null, "Excel", string.Empty, string.Empty);

                string filePath = _workbook.FullName;
                if (string.IsNullOrEmpty(filePath))
                    return (null, null, "Excel", _workbook.Name, string.Empty);

                var fileIdInfo = GetFileId(filePath);
                return (
                    fileIdInfo?.FileId,
                    fileIdInfo?.VolumeId,
                    "Excel",
                    _workbook.Name,
                    filePath
                );
            }
            catch (Exception ex)
            {
                Console.WriteLine($"파일 정보 가져오기 오류: {ex.Message}");
                return (null, null, "Excel", string.Empty, string.Empty);
            }
        }
    }
}

