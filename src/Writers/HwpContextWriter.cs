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

        public HwpContextWriter(ILogger<HwpContextWriter> logger)
        {
            _logger = logger;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        public void WriteContent(string content, Dictionary<string, object> styleAttributes)
        {
            try
            {
                _logger.LogInformation("한글 파일 쓰기 시작");
                var hwpProcesses = Process.GetProcessesByName("Hwp");
                
                if (hwpProcesses.Length == 0)
                {
                    throw new InvalidOperationException("한글(Hwp)이 실행 중이지 않습니다.");
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
                    throw new InvalidOperationException("활성 한글 창을 찾을 수 없습니다.");
                }

                // 한글 자동화 객체 가져오기
                var hwpApp = GetActiveObject("HwpFrame.HwpObject");
                if (hwpApp == null)
                {
                    throw new InvalidOperationException("한글 자동화 객체를 가져올 수 없습니다.");
                }

                // 현재 문서에 내용 삽입
                dynamic hwp = hwpApp;
                hwp.InsertText(content);

                _logger.LogInformation("한글 파일 쓰기 완료");
            }
            catch (Exception ex)
            {
                _logger.LogError($"한글 파일 쓰기 오류: {ex.Message}");
                throw;
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
