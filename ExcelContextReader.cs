using System.Collections.Generic;
using System.Windows.Automation;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.VisualBasic;
using System.Diagnostics;
using System.Threading;
using System.IO;
using Microsoft.Win32;

namespace overlay_gpt
{
    public class ExcelContextReader : BaseContextReader
    {
        private const int MAX_RETRY_COUNT = 3;
        private const int RETRY_DELAY_MS = 100;
        private Excel.Application? _excelApp;
        private Excel.Workbook? _workbook;
        private Excel.Worksheet? _worksheet;
        private Excel.Range? _selection;
        private static readonly string LogFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "excel_errors.log");

        private string GetExcelVersion()
        {
            try
            {
                // 64비트 Excel 버전 확인
                using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"))
                {
                    if (key != null)
                    {
                        var version = key.GetValue("VersionToReport") as string;
                        if (!string.IsNullOrEmpty(version))
                        {
                            return $"64비트 Excel 버전: {version}";
                        }
                    }
                }

                // 32비트 Excel 버전 확인
                using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Wow6432Node\Microsoft\Office\ClickToRun\Configuration"))
                {
                    if (key != null)
                    {
                        var version = key.GetValue("VersionToReport") as string;
                        if (!string.IsNullOrEmpty(version))
                        {
                            return $"32비트 Excel 버전: {version}";
                        }
                    }
                }

                // 레거시 설치 방식 확인
                using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\Excel\Addins"))
                {
                    if (key != null)
                    {
                        var version = key.GetValue("Version") as string;
                        if (!string.IsNullOrEmpty(version))
                        {
                            return $"레거시 Excel 버전: {version}";
                        }
                    }
                }

                return "Excel 버전을 찾을 수 없습니다.";
            }
            catch (Exception ex)
            {
                return $"Excel 버전 확인 중 오류 발생: {ex.Message}";
            }
        }

        private void LogToFile(string message)
        {
            try
            {
                string logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";
                File.AppendAllText(LogFilePath, logMessage + Environment.NewLine);
            }
            catch
            {
                // 로깅 실패 시 무시
            }
        }

        public override (string SelectedText, Dictionary<string, object> StyleAttributes) GetSelectedTextWithStyle()
        {
            var styleAttributes = new Dictionary<string, object>();
            string selectedText = string.Empty;

            if (!IsExcelRunning())
            {
                string errorMessage = "Excel이 실행 중이지 않습니다.";
                LogWindow.Instance.Log(errorMessage);
                LogToFile(errorMessage);
                return (string.Empty, styleAttributes);
            }

            for (int retryCount = 0; retryCount < MAX_RETRY_COUNT; retryCount++)
            {
                try
                {
                    if (InitializeExcel())
                    {
                        var result = GetSelectedTextWithStyleInternal();
                        if (result != null)
                        {
                            return result.Value;
                        }
                    }
                    else
                    {
                        string errorMessage = "Excel 연결에 실패했습니다.";
                        LogWindow.Instance.Log(errorMessage);
                        LogToFile(errorMessage);
                        Thread.Sleep(RETRY_DELAY_MS);
                    }
                }
                catch (System.Exception ex)
                {
                    string errorMessage = $"Excel 읽기 시도 {retryCount + 1} 실패: {ex.Message}";
                    LogWindow.Instance.Log(errorMessage);
                    LogToFile(errorMessage);
                    LogToFile($"스택 트레이스: {ex.StackTrace}");
                    Thread.Sleep(RETRY_DELAY_MS);
                }
                finally
                {
                    CleanupExcel();
                }
            }

            return (string.Empty, styleAttributes);
        }

        private bool InitializeExcel()
        {
            try
            {
                // Excel 버전 정보 로깅
                string excelVersion = GetExcelVersion();
                LogToFile($"설치된 Excel 정보: {excelVersion}");

                // 먼저 실행 중인 Excel 프로세스를 찾습니다
                var processes = Process.GetProcessesByName("EXCEL");
                if (processes.Length == 0)
                {
                    LogToFile("실행 중인 Excel 프로세스를 찾을 수 없습니다.");
                    return false;
                }

                // COM 객체를 통해 Excel 애플리케이션에 연결
                try
                {
                    Type? excelType = Type.GetTypeFromProgID("Excel.Application");
                    if (excelType == null)
                    {
                        LogToFile("Excel COM 객체를 찾을 수 없습니다.");
                        return false;
                    }

                    object? excelObj = null;
                    try
                    {
                        excelObj = Interaction.GetObject(null, "Excel.Application");
                    }
                    catch
                    {
                        // GetObject 실패 시 CreateInstance 시도
                        excelObj = Activator.CreateInstance(excelType);
                    }

                    if (excelObj == null)
                    {
                        LogToFile("Excel 애플리케이션을 생성할 수 없습니다.");
                        return false;
                    }

                    _excelApp = (Excel.Application)excelObj;
                    if (_excelApp == null)
                    {
                        LogToFile("Excel 애플리케이션에 연결할 수 없습니다.");
                        return false;
                    }
                }
                catch (Exception ex)
                {
                    LogToFile($"Excel COM 객체 연결 실패: {ex.Message}");
                    return false;
                }

                try
                {
                    _workbook = _excelApp.ActiveWorkbook;
                    if (_workbook == null)
                    {
                        LogToFile("활성 워크북을 찾을 수 없습니다.");
                        return false;
                    }

                    _worksheet = _excelApp.ActiveSheet as Excel.Worksheet;
                    if (_worksheet == null)
                    {
                        LogToFile("활성 워크시트를 찾을 수 없습니다.");
                        return false;
                    }

                    _selection = _excelApp.Selection as Excel.Range;
                    if (_selection == null)
                    {
                        LogToFile("선택된 범위를 찾을 수 없습니다.");
                        return false;
                    }
                }
                catch (Exception ex)
                {
                    LogToFile($"Excel 객체 접근 실패: {ex.Message}");
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                LogToFile($"Excel 초기화 중 오류 발생: {ex.Message}");
                LogToFile($"스택 트레이스: {ex.StackTrace}");
                return false;
            }
        }

        private void CleanupExcel()
        {
            try
            {
                if (_selection != null)
                {
                    Marshal.ReleaseComObject(_selection);
                    _selection = null;
                }
                if (_worksheet != null)
                {
                    Marshal.ReleaseComObject(_worksheet);
                    _worksheet = null;
                }
                if (_workbook != null)
                {
                    Marshal.ReleaseComObject(_workbook);
                    _workbook = null;
                }
                if (_excelApp != null)
                {
                    Marshal.ReleaseComObject(_excelApp);
                    _excelApp = null;
                }
            }
            catch (System.Exception ex)
            {
                LogWindow.Instance.Log($"Excel 리소스 해제 오류: {ex.Message}");
            }
        }

        private (string SelectedText, Dictionary<string, object> StyleAttributes)? GetSelectedTextWithStyleInternal()
        {
            if (_selection == null)
            {
                return null;
            }

            var styleAttributes = new Dictionary<string, object>();
            string selectedText = _selection.Text?.ToString() ?? string.Empty;

            if (_selection.Cells.Count > 0)
            {
                Excel.Range? firstCell = null;
                try
                {
                    firstCell = _selection.Cells[1, 1] as Excel.Range;
                    if (firstCell != null)
                    {
                        styleAttributes["FontName"] = firstCell.Font.Name ?? string.Empty;
                        styleAttributes["FontSize"] = firstCell.Font.Size;
                        styleAttributes["FontWeight"] = ((bool)firstCell.Font.Bold) ? 700 : 400;
                        styleAttributes["ForegroundColor"] = firstCell.Font.Color;
                        styleAttributes["BackgroundColor"] = firstCell.Interior.Color;
                    }
                }
                finally
                {
                    if (firstCell != null)
                    {
                        Marshal.ReleaseComObject(firstCell);
                    }
                }
            }

            return (selectedText, styleAttributes);
        }

        private bool IsExcelRunning()
        {
            try
            {
                var processes = Process.GetProcessesByName("EXCEL");
                return processes.Length > 0;
            }
            catch
            {
                return false;
            }
        }
    }
}
