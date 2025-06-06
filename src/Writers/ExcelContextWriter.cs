using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace overlay_gpt
{
    public class ExcelContextWriter : IContextWriter, IDisposable
    {
        private Microsoft.Office.Interop.Excel.Application? _excelApp;
        private Workbook? _workbook;

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

        /// <summary>
        /// Excel 파일을 열어서 _excelApp, _workbook에 설정합니다.
        /// </summary>
        public bool OpenFile(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"파일이 존재하지 않습니다: {filePath}");
                    return false;
                }

                _excelApp = (Microsoft.Office.Interop.Excel.Application)GetActiveObject("Excel.Application");
                _workbook = _excelApp.Workbooks.Open(filePath);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"파일 열기 오류: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// HTML 텍스트(테이블 태그 포함)를 받아서 임시 HTML 파일로 저장한 뒤,
        /// 해당 파일을 워크북으로 열어 복사 → Range.Copy(Destination) 방식으로
        /// 원본 시트에 삽입합니다.
        /// </summary>
        /// <param name="htmlText">
        /// “<table>…</table>” 단편 형태의 HTML이라 가정합니다.
        /// </param>
        /// <param name="lineNumber">
        /// 붙여넣기를 원하는 Excel 범위(예: "$A$1:$C$5" 등).
        /// </param>
        public bool ApplyTextWithStyle(string htmlText, string lineNumber)
        {
            if (_excelApp == null || _workbook == null)
            {
                Console.WriteLine("Excel 애플리케이션 또는 워크북이 초기화되지 않았습니다.");
                return false;
            }

            try
            {
                Console.WriteLine("\n=== Excel 데이터 적용(수정된 방식: Copy(Destination)) 시작 ===");
                Console.WriteLine($"위치: {lineNumber}");
                Console.WriteLine($"입력 HTML 단편 길이: {htmlText.Length}자");
                Console.WriteLine($"입력 HTML 단편 (처음 100자): {htmlText.Substring(0, Math.Min(100, htmlText.Length))}");

                // 1) HTML 단편(fragment)을 완전한 HTML 문서로 감싼다.
                string fullHtml = $@"<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <style>
        table {{ border-collapse: collapse; }}
        td {{ padding: 2px; font-family: Arial, sans-serif; font-size: 11pt; }}
    </style>
</head>
<body>
    {htmlText}
</body>
</html>";

                Console.WriteLine($"전체 HTML 길이: {fullHtml.Length}자");

                // 2) 임시 파일 경로를 얻고, UTF-8로 저장
                string tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".html");
                File.WriteAllText(tempFile, fullHtml, Encoding.UTF8);
                File.SetAttributes(tempFile, FileAttributes.Hidden);
                Console.WriteLine($"임시 HTML 파일 생성: {tempFile}");

                // 3) 임시 HTML 파일을 Excel 워크북으로 연다.
                Workbook tempWb;
                Worksheet tempSheet;
                try
                {
                    tempWb = _excelApp.Workbooks.Open(
                        tempFile,
                        ReadOnly: true,
                        Editable: false,
                        Local: true
                    );
                    tempSheet = (Worksheet)tempWb.ActiveSheet;
                    Console.WriteLine($"임시 워크북 열기 성공: 시트 이름 = {tempSheet.Name}");
                }
                catch (Exception openTempEx)
                {
                    Console.WriteLine($"임시 HTML 워크북 열기 실패: {openTempEx.Message}");
                    if (File.Exists(tempFile))
                        File.Delete(tempFile);
                    return false;
                }

                // 4) 임시 시트의 UsedRange 전체를 가져와서, Destination으로 Copy
                Range copyRange = tempSheet.UsedRange;
                Console.WriteLine($"임시 시트 UsedRange: {copyRange.Address}");

                // 5) 원본 워크북의 대상 시트를 가져오고, 대상 범위를 Clear
                Worksheet origSheet = (Worksheet)_workbook.ActiveSheet;
                origSheet.Activate();

                Range targetRange = origSheet.Range[lineNumber];
                if (targetRange == null)
                {
                    Console.WriteLine($"원본 워크북의 대상 범위({lineNumber})를 찾을 수 없습니다.");
                    tempWb.Close(false);
                    if (File.Exists(tempFile))
                        File.Delete(tempFile);
                    return false;
                }

                // (Optional) 원본 범위 초기화: 값만 지우고 서식은 남기는 경우라면 ClearContents(),
                // 서식까지 지우려면 Clear() 사용.
                Console.WriteLine($"원본 워크북의 대상 범위: {targetRange.Address}");
                targetRange.Clear();
                Console.WriteLine("원본 범위 내용 및 서식 지움 완료.");

                // 6) copyRange를 targetRange를 시작점으로 삼아서 붙여넣기
                //    이때 copyRange.Copy(Destination) 메서드를 사용하면 클립보드 불필요
                Range destinationCell = origSheet.Range[lineNumber].Cells[1, 1];
                // copyRange 크기가 예: A1:C5라면, Destination 셀부터 A1상의 위치에 동일 크기로 복사된다.
                copyRange.Copy(destinationCell);
                Console.WriteLine($"copyRange.Copy(dest) 호출 완료: dest = {destinationCell.Address}");

                // 7) 임시 워크북 닫고, COM 개체 해제 및 파일 삭제
                tempWb.Close(false);
                Marshal.ReleaseComObject(tempSheet);
                Marshal.ReleaseComObject(tempWb);
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                    Console.WriteLine("임시 HTML 파일 삭제 완료.");
                }

                Console.WriteLine("=== Excel 데이터 적용(수정된 방식: Copy(Destination)) 완료 ===\n");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Excel 데이터 적용 오류: {ex.Message}");
                Console.WriteLine($"스택 트레이스: {ex.StackTrace}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"내부 예외: {ex.InnerException.Message}");
                    Console.WriteLine($"내부 예외 스택 트레이스: {ex.InnerException.StackTrace}");
                }
                return false;
            }
        }

        /// <summary>
        /// 현재 열려 있는 워크북의 파일 정보를 반환합니다.
        /// </summary>
        public (ulong? FileId, uint? VolumeId, string FileType, string FileName, string FilePath) GetFileInfo()
        {
            try
            {
                if (_workbook == null)
                    return (null, null, "Excel", string.Empty, string.Empty);

                string filePath = _workbook.FullName;
                string fileName = _workbook.Name;
                return (null, null, "Excel", fileName, filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Excel 파일 정보 가져오기 오류: {ex.Message}");
                return (null, null, "Excel", string.Empty, string.Empty);
            }
        }

        /// <summary>
        /// 리소스를 해제합니다. 반드시 사용 후 호출해야 합니다.
        /// </summary>
        public void Dispose()
        {
            try
            {
                if (_workbook != null)
                {
                    _workbook.Close(false);
                    Marshal.ReleaseComObject(_workbook);
                    _workbook = null;
                }

                if (_excelApp != null)
                {
                    _excelApp.Quit();
                    Marshal.ReleaseComObject(_excelApp);
                    _excelApp = null;
                }
            }
            catch
            {
                // COM 해제 중 예외 발생 시 무시
            }
        }
    }
}
