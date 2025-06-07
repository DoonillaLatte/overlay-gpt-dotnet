using System;
using System.Collections.Generic;
using System.Windows.Automation;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using System.Xml;
using Microsoft.Extensions.Logging;
using System.IO.Compression;

namespace overlay_gpt
{
    public class HwpContextWriter : IContextWriter
    {
        private readonly ILogger<HwpContextWriter> _logger;
        private string? _currentFilePath;
        private bool _isTargetProg;

        public bool IsTargetProg
        {
            get => _isTargetProg;
            set => _isTargetProg = value;
        }

        public HwpContextWriter(ILogger<HwpContextWriter> logger, bool isTargetProg = false)
        {
            _logger = logger;
            _isTargetProg = isTargetProg;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

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
                    _logger.LogError($"파일이 존재하지 않습니다: {filePath}");
                    return false;
                }

                try
                {
                    _logger.LogInformation("기존 한글 애플리케이션 찾기 시도...");
                    var hwpApp = GetActiveObject("HwpFrame.HwpObject");
                    
                    if(hwpApp != null)
                    {
                        _logger.LogInformation("기존 한글 애플리케이션 찾음");
                        _currentFilePath = filePath;
                        return true;
                    }
                    else
                    {
                        _logger.LogError("기존 한글 애플리케이션을 찾을 수 없습니다.");
                        return false;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError($"한글 애플리케이션이 실행 중이지 않습니다: {ex.Message}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"파일 열기 오류: {ex.Message}");
                return false;
            }
        }

        public bool ApplyTextWithStyle(string text, string lineNumber)
        {
            try
            {
                _logger.LogInformation("한글 파일 쓰기 시작");
                var hwpProcesses = Process.GetProcessesByName("Hwp");
                
                if (hwpProcesses.Length == 0)
                {
                    _logger.LogError("한글(Hwp)이 실행 중이지 않습니다.");
                    return false;
                }

                Process? activeHwpProcess = null;
                foreach (var process in hwpProcesses)
                {
                    if (process.MainWindowHandle != IntPtr.Zero && process.MainWindowHandle == GetForegroundWindow())
                    {
                        activeHwpProcess = process;
                        break;
                    }
                }

                if (activeHwpProcess == null)
                {
                    _logger.LogError("활성 한글 창을 찾을 수 없습니다.");
                    return false;
                }

                // 한글 자동화 객체 가져오기
                var hwpApp = GetActiveObject("HwpFrame.HwpObject");
                if (hwpApp == null)
                {
                    _logger.LogError("한글 자동화 객체를 가져올 수 없습니다.");
                    return false;
                }

                // 현재 문서에 내용 삽입
                dynamic hwp = hwpApp;
                hwp.InsertText(text);

                _logger.LogInformation("한글 파일 쓰기 완료");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError($"한글 파일 쓰기 오류: {ex.Message}");
                return false;
            }
        }

        public (ulong? FileId, uint? VolumeId, string FileType, string FileName, string FilePath) GetFileInfo()
        {
            if (string.IsNullOrEmpty(_currentFilePath))
            {
                return (null, null, "Hwp", string.Empty, string.Empty);
            }

            return (
                null, // FileId는 현재 구현되지 않음
                null, // VolumeId는 현재 구현되지 않음
                "Hwp",
                Path.GetFileName(_currentFilePath),
                _currentFilePath
            );
        }
    }
}
