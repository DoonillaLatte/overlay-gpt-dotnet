using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using HtmlAgilityPack;

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
                Console.WriteLine("Excel COM 객체 생성 시도...");
                _excelApp = new Application();
                _excelApp.Visible = false; // 백그라운드에서 실행
                Console.WriteLine("Excel COM 객체 생성 성공");

                Console.WriteLine($"파일 열기 시도: {filePath}");
                _workbook = _excelApp.Workbooks.Open(filePath);
                _worksheet = _workbook.ActiveSheet;
                Console.WriteLine("파일 열기 성공");

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Excel 파일 열기 오류 발생: {ex.Message}");
                Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                return false;
            }
        }

        public bool ApplyTextWithStyle(string text, string lineNumber)
        {
            try
            {
                if (_excelApp == null || _workbook == null || _worksheet == null)
                {
                    Console.WriteLine("Excel 애플리케이션이 초기화되지 않았습니다.");
                    return false;
                }

                // 라인 번호 파싱 (예: "R1C1-R1C1")
                var lineNumbers = lineNumber.Split('-');
                if (lineNumbers.Length != 2)
                {
                    Console.WriteLine("잘못된 라인 번호 형식입니다.");
                    return false;
                }

                // 시작 셀과 끝 셀 파싱
                var startCell = lineNumbers[0].Replace("R", "").Replace("C", ",").Split(',');
                var endCell = lineNumbers[1].Replace("R", "").Replace("C", ",").Split(',');

                if (startCell.Length != 2 || endCell.Length != 2)
                {
                    Console.WriteLine("잘못된 셀 형식입니다.");
                    return false;
                }

                int startRow = int.Parse(startCell[0]);
                int startCol = int.Parse(startCell[1]);
                int endRow = int.Parse(endCell[0]);
                int endCol = int.Parse(endCell[1]);

                // 선택된 범위 설정
                var range = _worksheet.Range[
                    _worksheet.Cells[startRow, startCol],
                    _worksheet.Cells[endRow, endCol]
                ];

                // HTML 태그 처리
                var htmlDoc = new HtmlDocument();
                htmlDoc.LoadHtml(text);

                // 텍스트와 스타일 적용
                foreach (var node in htmlDoc.DocumentNode.ChildNodes)
                {
                    if (node.NodeType == HtmlNodeType.Text)
                    {
                        range.Text = node.InnerText;
                    }
                    else
                    {
                        var style = node.GetAttributeValue("style", "");
                        var font = range.Font;

                        // 스타일 속성 파싱
                        var styleAttributes = style.Split(';')
                            .Select(s => s.Trim().Split(':'))
                            .Where(p => p.Length == 2)
                            .ToDictionary(p => p[0].Trim(), p => p[1].Trim());

                        // 폰트 패밀리
                        if (styleAttributes.TryGetValue("font-family", out var fontFamily))
                        {
                            font.Name = fontFamily.Trim('\'');
                        }

                        // 폰트 크기
                        if (styleAttributes.TryGetValue("font-size", out var fontSize))
                        {
                            if (fontSize.EndsWith("pt"))
                            {
                                font.Size = float.Parse(fontSize.Replace("pt", ""));
                            }
                        }

                        // 색상
                        if (styleAttributes.TryGetValue("color", out var color))
                        {
                            if (color.StartsWith("#"))
                            {
                                var rgb = int.Parse(color.Substring(1), System.Globalization.NumberStyles.HexNumber);
                                font.Color = rgb;
                            }
                        }

                        // 배경색
                        if (styleAttributes.TryGetValue("background-color", out var bgColor))
                        {
                            if (bgColor.StartsWith("#"))
                            {
                                var rgb = int.Parse(bgColor.Substring(1), System.Globalization.NumberStyles.HexNumber);
                                range.Interior.Color = rgb;
                            }
                        }

                        // 굵게
                        if (node.Name == "b" || node.Name == "strong")
                        {
                            font.Bold = true;
                        }

                        // 기울임
                        if (node.Name == "i" || node.Name == "em")
                        {
                            font.Italic = true;
                        }

                        // 밑줄
                        if (node.Name == "u")
                        {
                            font.Underline = XlUnderlineStyle.xlUnderlineStyleSingle;
                        }

                        // 취소선
                        if (node.Name == "s" || node.Name == "strike")
                        {
                            font.Strikethrough = true;
                        }

                        // 텍스트 정렬
                        if (styleAttributes.TryGetValue("text-align", out var textAlign))
                        {
                            range.HorizontalAlignment = textAlign switch
                            {
                                "center" => XlHAlign.xlHAlignCenter,
                                "right" => XlHAlign.xlHAlignRight,
                                "justify" => XlHAlign.xlHAlignJustify,
                                _ => XlHAlign.xlHAlignLeft
                            };
                        }

                        if (styleAttributes.TryGetValue("vertical-align", out var verticalAlign))
                        {
                            range.VerticalAlignment = verticalAlign switch
                            {
                                "top" => XlVAlign.xlVAlignTop,
                                "middle" => XlVAlign.xlVAlignCenter,
                                _ => XlVAlign.xlVAlignBottom
                            };
                        }

                        range.Text = node.InnerText;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"텍스트 적용 중 오류 발생: {ex.Message}");
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
                
                Console.WriteLine($"Excel 문서 정보:");
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
    }
} 