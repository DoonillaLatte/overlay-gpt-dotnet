using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using HtmlAgilityPack;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace overlay_gpt
{
    public class ExcelContextWriter : IContextWriter
    {
        private Application? _excelApp;
        private Workbook? _workbook;
        private Worksheet? _worksheet;

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
                Console.WriteLine("기존 Excel 프로세스 확인 중...");
                try
                {
                    _excelApp = (Application)GetActiveObject("Excel.Application");
                    Console.WriteLine("기존 Excel 프로세스 발견");

                    // 이미 열려있는 문서 확인
                    foreach (Workbook wb in _excelApp.Workbooks)
                    {
                        try
                        {
                            if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("파일이 이미 열려있습니다.");
                                _workbook = wb;
                                _worksheet = _excelApp.ActiveSheet as Worksheet;
                                return true;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"워크북 확인 중 오류 발생: {ex.Message}");
                            continue;
                        }
                    }
                }
                catch
                {
                    Console.WriteLine("새로운 Excel COM 객체 생성 시도...");
                    _excelApp = new Application();
                    _excelApp.Visible = false; // 백그라운드에서 실행
                    Console.WriteLine("새로운 Excel COM 객체 생성 성공");
                }

                Console.WriteLine($"파일 열기 시도: {filePath}");
                _workbook = _excelApp.Workbooks.Open(filePath);
                _worksheet = _excelApp.ActiveSheet as Worksheet;
                Console.WriteLine("파일 열기 성공");

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Excel 파일 열기 오류 발생: {ex.Message}");
                Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                
                // 오류 발생 시 COM 객체 정리
                if (_worksheet != null)
                {
                    try { Marshal.ReleaseComObject(_worksheet); } catch { }
                    _worksheet = null;
                }
                if (_workbook != null)
                {
                    try { Marshal.ReleaseComObject(_workbook); } catch { }
                    _workbook = null;
                }
                if (_excelApp != null)
                {
                    try { Marshal.ReleaseComObject(_excelApp); } catch { }
                    _excelApp = null;
                }
                
                return false;
            }
        }

        public bool ApplyTextWithStyle(string text, string lineNumber)
        {
            try
            {
                if (_excelApp == null || _worksheet == null)
                {
                    Console.WriteLine("Excel 애플리케이션이 초기화되지 않았습니다.");
                    return false;
                }

                Console.WriteLine($"텍스트 적용 시작 - 라인 번호: {lineNumber}");
                Console.WriteLine($"적용할 텍스트: {text}");

                // 라인 번호 파싱 (예: "R1C1-R2C2")
                var lineNumbers = lineNumber.Split('-');
                if (lineNumbers.Length != 2)
                {
                    Console.WriteLine("잘못된 라인 번호 형식입니다.");
                    return false;
                }

                // 시작 셀과 끝 셀의 행/열 번호 추출
                var startCell = lineNumbers[0].Substring(1).Split('C');
                var endCell = lineNumbers[1].Substring(1).Split('C');

                int startRow = int.Parse(startCell[0]);
                int startCol = int.Parse(startCell[1]);
                int endRow = int.Parse(endCell[0]);
                int endCol = int.Parse(endCell[1]);

                Console.WriteLine($"시작 셀: R{startRow}C{startCol}, 종료 셀: R{endRow}C{endCol}");

                // 선택된 범위 설정
                Range selectedRange = null;
                try
                {
                    Console.WriteLine("셀 범위 설정 시도...");
                    selectedRange = _worksheet.Range[
                        _worksheet.Cells[startRow, startCol],
                        _worksheet.Cells[endRow, endCol]
                    ];
                    Console.WriteLine("셀 범위 설정 성공");

                    // 선택된 범위의 내용과 스타일 모두 지우기
                    Console.WriteLine("선택된 범위의 내용과 스타일 지우기...");
                    selectedRange.Clear();
                    Console.WriteLine("선택된 범위 초기화 완료");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"셀 범위 설정 중 오류 발생: {ex.Message}");
                    return false;
                }

                if (selectedRange == null)
                {
                    Console.WriteLine("셀 범위를 설정할 수 없습니다.");
                    return false;
                }

                // HTML 태그 처리
                Console.WriteLine("HTML 파싱 시작...");
                var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.LoadHtml(text);
                Console.WriteLine($"HTML 노드 수: {htmlDoc.DocumentNode.ChildNodes.Count}");

                // 테이블 구조 파싱
                var tableNode = htmlDoc.DocumentNode.SelectSingleNode("//table");
                if (tableNode != null)
                {
                    var rows = tableNode.SelectNodes(".//tr");
                    if (rows != null)
                    {
                        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
                        {
                            var cells = rows[rowIndex].SelectNodes(".//td");
                            if (cells != null)
                            {
                                for (int colIndex = 0; colIndex < cells.Count; colIndex++)
                                {
                                    var cell = cells[colIndex];
                                    var cellRange = _worksheet.Range[
                                        _worksheet.Cells[startRow + rowIndex, startCol + colIndex],
                                        _worksheet.Cells[startRow + rowIndex, startCol + colIndex]
                                    ];

                                    try
                                    {
                                        Console.WriteLine($"셀 처리 시작 - 행: {rowIndex + 1}, 열: {colIndex + 1}");
                                        
                                        // 텍스트 적용
                                        cellRange.Value2 = cell.InnerText.Trim();
                                        
                                        // 스타일 적용
                                        var style = cell.GetAttributeValue("style", "");
                                        Console.WriteLine($"셀 스타일: {style}");
                                        
                                        var styleAttributes = style.Split(';')
                                            .Select(s => s.Trim().Split(':'))
                                            .Where(p => p.Length == 2)
                                            .ToDictionary(p => p[0].Trim(), p => p[1].Trim());

                                        // 폰트 설정
                                        var font = cellRange.Font;

                                        // 배경색
                                        if (styleAttributes.TryGetValue("background-color", out var bgColor))
                                        {
                                            if (bgColor.StartsWith("#"))
                                            {
                                                Console.WriteLine($"배경색 설정: {bgColor}");
                                                var rgb = int.Parse(bgColor.Substring(1), System.Globalization.NumberStyles.HexNumber);
                                                // BGR로 변환 (Excel은 BGR 형식 사용)
                                                int b = (rgb >> 16) & 0xFF;
                                                int g = (rgb >> 8) & 0xFF;
                                                int r = rgb & 0xFF;
                                                int bgr = (r << 16) | (g << 8) | b;
                                                cellRange.Interior.Color = bgr;
                                            }
                                        }

                                        // 텍스트 색상
                                        if (styleAttributes.TryGetValue("color", out var color))
                                        {
                                            if (color.StartsWith("#"))
                                            {
                                                Console.WriteLine($"텍스트 색상 설정: {color}");
                                                var rgb = int.Parse(color.Substring(1), System.Globalization.NumberStyles.HexNumber);
                                                // BGR로 변환 (Excel은 BGR 형식 사용)
                                                int b = (rgb >> 16) & 0xFF;
                                                int g = (rgb >> 8) & 0xFF;
                                                int r = rgb & 0xFF;
                                                int bgr = (r << 16) | (g << 8) | b;
                                                font.Color = bgr;
                                            }
                                        }

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

                                        // 굵게
                                        if (styleAttributes.TryGetValue("font-weight", out var fontWeight) && fontWeight == "bold")
                                        {
                                            Console.WriteLine("굵게 스타일 적용");
                                            font.Bold = -1;
                                        }

                                        // 테두리
                                        var borderPositions = new Dictionary<string, (string style, XlBordersIndex index)>
                                        {
                                            { "border-top", (styleAttributes.TryGetValue("border-top", out var top) ? top.ToString() : "", XlBordersIndex.xlEdgeTop) },
                                            { "border-right", (styleAttributes.TryGetValue("border-right", out var right) ? right.ToString() : "", XlBordersIndex.xlEdgeRight) },
                                            { "border-bottom", (styleAttributes.TryGetValue("border-bottom", out var bottom) ? bottom.ToString() : "", XlBordersIndex.xlEdgeBottom) },
                                            { "border-left", (styleAttributes.TryGetValue("border-left", out var left) ? left.ToString() : "", XlBordersIndex.xlEdgeLeft) }
                                        };

                                        foreach (var position in borderPositions)
                                        {
                                            if (!string.IsNullOrEmpty(position.Value.style) && position.Value.style.Contains("solid"))
                                            {
                                                Console.WriteLine($"{position.Key} 스타일 적용");
                                                var border = cellRange.Borders[position.Value.index];
                                                border.LineStyle = XlLineStyle.xlContinuous;
                                                
                                                // 테두리 두께 설정
                                                if (position.Value.style.Contains("2px"))
                                                {
                                                    border.Weight = XlBorderWeight.xlMedium;
                                                }
                                                else if (position.Value.style.Contains("3px"))
                                                {
                                                    border.Weight = XlBorderWeight.xlThick;
                                                }
                                                else
                                                {
                                                    border.Weight = XlBorderWeight.xlThin;
                                                }
                                                
                                                border.Color = 0; // 검은색
                                            }
                                        }

                                        Console.WriteLine($"셀 처리 완료 - 행: {rowIndex + 1}, 열: {colIndex + 1}");
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"셀 처리 중 오류 발생: {ex.Message}");
                                        Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                                    }
                                    finally
                                    {
                                        if (cellRange != null)
                                        {
                                            try { Marshal.ReleaseComObject(cellRange); } catch { }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine("테이블 구조를 찾을 수 없습니다.");
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
            Application? tempExcelApp = null;
            Workbook? tempWorkbook = null;
            
            try
            {
                Console.WriteLine("Excel COM 객체 가져오기 시도...");
                tempExcelApp = (Application)GetActiveObject("Excel.Application");
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
                
                Console.WriteLine($"Excel 워크북 정보:");
                Console.WriteLine($"- 파일 경로: {filePath}");
                Console.WriteLine($"- 파일 이름: {fileName}");
                
                if (string.IsNullOrEmpty(filePath))
                {
                    Console.WriteLine("파일 경로가 비어있습니다.");
                    return (null, null, "Excel", fileName, string.Empty);
                }
                
                return (null, null, "Excel", fileName, filePath);
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

        public void Dispose()
        {
            if (_worksheet != null)
            {
                try { Marshal.ReleaseComObject(_worksheet); } catch { }
                _worksheet = null;
            }
            if (_workbook != null)
            {
                try { Marshal.ReleaseComObject(_workbook); } catch { }
                _workbook = null;
            }
            if (_excelApp != null)
            {
                try { Marshal.ReleaseComObject(_excelApp); } catch { }
                _excelApp = null;
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
