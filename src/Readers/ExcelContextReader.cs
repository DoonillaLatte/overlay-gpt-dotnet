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

        public override (string SelectedText, Dictionary<string, object> StyleAttributes) GetSelectedTextWithStyle()
        {
            try
            {
                Console.WriteLine("Excel 데이터 읽기 시작...");

                // 실행 중인 Excel 애플리케이션 가져오기
                Console.WriteLine("실행 중인 Excel 애플리케이션 가져오기 시도...");
                var excelProcesses = Process.GetProcessesByName("EXCEL");
                if (excelProcesses.Length == 0)
                {
                    Console.WriteLine("실행 중인 Excel 애플리케이션을 찾을 수 없습니다.");
                    return (string.Empty, new Dictionary<string, object>());
                }

                // 활성화된 Excel 창 찾기
                Process? activeExcelProcess = null;
                foreach (var process in excelProcesses)
                {
                    if (process.MainWindowHandle != IntPtr.Zero && process.MainWindowTitle.Length > 0)
                    {
                        Console.WriteLine($"Excel 프로세스 정보:");
                        Console.WriteLine($"- 프로세스 ID: {process.Id}");
                        Console.WriteLine($"- 시작 시간: {process.StartTime}");
                        Console.WriteLine($"- 메모리 사용량: {process.WorkingSet64 / 1024 / 1024} MB");
                        Console.WriteLine($"- 창 제목: {process.MainWindowTitle}");
                        Console.WriteLine($"- 실행 경로: {process.MainModule?.FileName}");
                        Console.WriteLine("------------------------");

                        // 현재 활성화된 창인지 확인
                        if (process.MainWindowHandle == GetForegroundWindow())
                        {
                            activeExcelProcess = process;
                            Console.WriteLine("이 Excel 창이 현재 활성화되어 있습니다.");
                        }
                    }
                }

                if (activeExcelProcess == null)
                {
                    Console.WriteLine("활성화된 Excel 창을 찾을 수 없습니다. Excel 창을 선택해주세요.");
                    return (string.Empty, new Dictionary<string, object>());
                }

                // COM을 통해 실행 중인 Excel 인스턴스에 연결
                try
                {
                    // 실행 중인 Excel 인스턴스에 직접 연결
                    _excelApp = (Application)GetActiveObject("Excel.Application");
                    
                    // 실행 중인 모든 Excel 인스턴스 가져오기
                    var runningInstances = _excelApp.Workbooks;
                    Console.WriteLine($"실행 중인 Excel 인스턴스 수: {runningInstances.Count}");
                    foreach (Workbook wb in runningInstances)
                    {
                        Console.WriteLine($"Excel 인스턴스 정보: {wb.FullName}");
                    }
                    
                    // 활성화된 Excel 인스턴스 찾기
                    _workbook = _excelApp.ActiveWorkbook;
                    if (_workbook != null)
                    {
                        Console.WriteLine($"활성 워크북 찾음: {_workbook.FullName}");
                    }

                    if (_excelApp == null)
                    {
                        Console.WriteLine("실행 중인 Excel 애플리케이션에 연결할 수 없습니다.");
                        return (string.Empty, new Dictionary<string, object>());
                    }
                    Console.WriteLine("실행 중인 Excel 애플리케이션 연결 성공");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Excel COM 연결 오류: {ex.Message}");
                    return (string.Empty, new Dictionary<string, object>());
                }

                // 현재 활성화된 워크북과 워크시트 가져오기
                Console.WriteLine("활성 워크북 가져오기 시도...");
                if (_workbook == null)
                {
                    _workbook = _excelApp.ActiveWorkbook;
                }
                
                if (_workbook != null)
                {
                    string filePath = _workbook.FullName;
                    Console.WriteLine($"현재 편집 중인 Excel 파일 정보:");
                    Console.WriteLine($"- 파일 이름: {_workbook.Name}");
                    Console.WriteLine($"- 전체 경로: {filePath}");
                    Console.WriteLine($"- 저장 여부: {(_workbook.Saved ? "저장됨" : "저장되지 않음")}");
                    Console.WriteLine($"- 읽기 전용: {(_workbook.ReadOnly ? "예" : "아니오")}");

                    // NTFS 파일 ID 가져오기
                    var fileIdInfo = GetFileId(filePath);
                    if (fileIdInfo.HasValue)
                    {
                        Console.WriteLine($"- NTFS 파일 ID: {fileIdInfo.Value.FileId}");
                        Console.WriteLine($"- 볼륨 ID: {fileIdInfo.Value.VolumeId}");
                    }
                    Console.WriteLine("------------------------");
                }
                Console.WriteLine("활성 워크시트 가져오기 시도...");
                _worksheet = _excelApp.ActiveSheet;

                if (_worksheet == null)
                {
                    Console.WriteLine("활성 워크시트를 찾을 수 없습니다.");
                    return (string.Empty, new Dictionary<string, object>());
                }
                Console.WriteLine("활성 워크시트 찾음");

                // 현재 선택된 범위 가져오기
                Console.WriteLine("선택된 범위 가져오기 시도...");
                Microsoft.Office.Interop.Excel.Range selectedRange = _excelApp.Selection;
                
                // 선택된 범위의 상세 정보 출력
                Console.WriteLine($"선택된 범위 정보:");
                Console.WriteLine($"- 주소: {selectedRange.Address}");
                Console.WriteLine($"- 행 수: {selectedRange.Rows.Count}");
                Console.WriteLine($"- 열 수: {selectedRange.Columns.Count}");

                // 2차원 배열 데이터 처리
                object[,] values = selectedRange.Value2 as object[,];
                StringBuilder selectedText = new StringBuilder();
                if (values != null)
                {
                    selectedText.Append("<table style='border-collapse:collapse'>");
                    for (int i = 1; i <= values.GetLength(0); i++)
                    {
                        selectedText.Append("<tr>");
                        for (int j = 1; j <= values.GetLength(1); j++)
                        {
                            var cell = selectedRange.Cells[i, j];
                            var cellStyle = new Dictionary<string, object>
                            {
                                ["FontName"] = cell.Font.Name,
                                ["FontSize"] = cell.Font.Size,
                                ["FontWeight"] = cell.Font.Bold ? "Bold" : "Normal",
                                ["FontItalic"] = cell.Font.Italic,
                                ["FontStrikethrough"] = cell.Font.Strikethrough,
                                ["FontSuperscript"] = cell.Font.Superscript,
                                ["FontSubscript"] = cell.Font.Subscript,
                                ["FontShadow"] = cell.Font.Shadow,
                                ["ForegroundColor"] = cell.Font.Color,
                                ["BackgroundColor"] = cell.Interior.Color,
                                ["UnderlineStyle"] = cell.Font.Underline == 2 ? "Single" : "None",
                                ["HorizontalAlignment"] = cell.HorizontalAlignment.ToString(),
                                ["VerticalAlignment"] = cell.VerticalAlignment.ToString(),
                                ["IndentLevel"] = cell.IndentLevel,
                                ["BorderStyle"] = new Dictionary<string, object>
                                {
                                    ["top"] = new Dictionary<string, object>
                                    {
                                        ["LineStyle"] = SafeGetInt(cell.Borders[XlBordersIndex.xlEdgeTop].LineStyle),
                                        ["Color"] = ConvertBGRToRGB(SafeGetInt(cell.Borders[XlBordersIndex.xlEdgeTop].Color))
                                    },
                                    ["right"] = new Dictionary<string, object>
                                    {
                                        ["LineStyle"] = SafeGetInt(cell.Borders[XlBordersIndex.xlEdgeRight].LineStyle),
                                        ["Color"] = ConvertBGRToRGB(SafeGetInt(cell.Borders[XlBordersIndex.xlEdgeRight].Color))
                                    },
                                    ["bottom"] = new Dictionary<string, object>
                                    {
                                        ["LineStyle"] = SafeGetInt(cell.Borders[XlBordersIndex.xlEdgeBottom].LineStyle),
                                        ["Color"] = ConvertBGRToRGB(SafeGetInt(cell.Borders[XlBordersIndex.xlEdgeBottom].Color))
                                    },
                                    ["left"] = new Dictionary<string, object>
                                    {
                                        ["LineStyle"] = SafeGetInt(cell.Borders[XlBordersIndex.xlEdgeLeft].LineStyle),
                                        ["Color"] = ConvertBGRToRGB(SafeGetInt(cell.Borders[XlBordersIndex.xlEdgeLeft].Color))
                                    }
                                }
                            };
                            // 디버깅: 셀의 border 정보 출력
                            var borderDict = (Dictionary<string, object>)cellStyle["BorderStyle"];
                            string cellValue = values[i, j]?.ToString() ?? "";
                            string cellText = GetStyledText(cellValue, cellStyle);
                            string cellStyleStr = GetCellStyleString(cellStyle);
                            selectedText.Append($"<td style='{cellStyleStr}'>{cellText}</td>");
                        }
                        selectedText.Append("</tr>");
                    }
                    selectedText.Append("</table>");
                }
                
                Console.WriteLine($"- 처리된 텍스트:\n{selectedText}");
                Console.WriteLine("------------------------");

                // 스타일 속성 가져오기
                Console.WriteLine("스타일 속성 가져오기 시도...");
                var styleAttributes = new Dictionary<string, object>
                {
                    ["FontName"] = selectedRange.Font.Name ?? "Calibri",
                    ["FontSize"] = selectedRange.Font.Size ?? 11.0,
                    ["FontWeight"] = SafeGetBoolean(selectedRange.Font.Bold) ? "Bold" : "Normal",
                    ["FontItalic"] = SafeGetBoolean(selectedRange.Font.Italic),
                    ["FontStrikethrough"] = SafeGetBoolean(selectedRange.Font.Strikethrough),
                    ["FontSuperscript"] = SafeGetBoolean(selectedRange.Font.Superscript),
                    ["FontSubscript"] = SafeGetBoolean(selectedRange.Font.Subscript),
                    ["FontShadow"] = SafeGetBoolean(selectedRange.Font.Shadow),
                    ["ForegroundColor"] = ConvertBGRToRGB(SafeGetInt(selectedRange.Font.Color, 0)),
                    ["BackgroundColor"] = ConvertBGRToRGB(SafeGetInt(selectedRange.Interior.Color, 16777215)),
                    ["UnderlineStyle"] = SafeGetInt(selectedRange.Font.Underline, 0) == 2 ? "Single" : "None",
                    ["HorizontalAlignment"] = selectedRange.HorizontalAlignment.ToString() ?? "Left",
                    ["VerticalAlignment"] = selectedRange.VerticalAlignment.ToString() ?? "Bottom",
                    ["IndentLevel"] = SafeGetInt(selectedRange.IndentLevel, 0)
                };
                Console.WriteLine("스타일 속성 가져오기 완료");

                return (selectedText.ToString(), styleAttributes);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Excel 데이터 읽기 오류 발생: {ex.Message}");
                Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                LogWindow.Instance.Log($"Excel 데이터 읽기 오류: {ex.Message}");
                return (string.Empty, new Dictionary<string, object>());
            }
            finally
            {
                Console.WriteLine("COM 객체 해제 시작...");
                if (_worksheet != null) Marshal.ReleaseComObject(_worksheet);
                if (_workbook != null) Marshal.ReleaseComObject(_workbook);
                if (_excelApp != null) Marshal.ReleaseComObject(_excelApp);
                Console.WriteLine("COM 객체 해제 완료");
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

