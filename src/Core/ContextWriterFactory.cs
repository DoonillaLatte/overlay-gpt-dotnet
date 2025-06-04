using System;
using System.Diagnostics;

namespace overlay_gpt
{
    public class ContextWriterFactory
    {
        public static IContextWriter CreateWriter(string fileType)
        {
            switch (fileType)
            {
                case "Word":
                    return new WordContextWriter();
                case "Excel":
                    return new ExcelContextWriter();
                case "PowerPoint":
                    return new PPTContextWriter();
                default:
                    throw new NotSupportedException($"지원하지 않는 프로그램입니다: {fileType}");
            }
        }

        private static Process? GetForegroundProcess()
        {
            try
            {
                IntPtr hwnd = GetForegroundWindow();
                if (hwnd == IntPtr.Zero)
                    return null;

                uint processId;
                GetWindowThreadProcessId(hwnd, out processId);
                return Process.GetProcessById((int)processId);
            }
            catch
            {
                return null;
            }
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    }
}
