using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Automation;
using System.Runtime.InteropServices;
using System.Threading;
using System.Security.Principal;
using System.Windows.Forms;    // SendKeys
using Application = System.Windows.Application;
using Clipboard = System.Windows.Clipboard;  // 명시적 Clipboard 지정

namespace overlay_gpt
{
    public class ClipboardContextReader : BaseContextReader
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError=true)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        private const int MAX_RETRIES    = 5;
        private const int RETRY_DELAY_MS = 200;

        private void CopyToClipboard()
        {
            Console.WriteLine("=== CopyToClipboard 시작 ===");
            
            // 1. 현재 포커스된 요소에서 시작
            var focused = AutomationElement.FocusedElement;
            if (focused == null)
            {
                Console.WriteLine("[오류] 포커스된 요소가 없습니다.");
                return;
            }

            // 2. 부모 윈도우 찾기
            AutomationElement parent = focused;
            IntPtr hwnd = IntPtr.Zero;

            while (parent != null)
            {
                hwnd = new IntPtr(parent.Current.NativeWindowHandle);
                if (hwnd != IntPtr.Zero)
                {
                    Console.WriteLine($"[정보] 유효한 윈도우 핸들 발견: {hwnd}");
                    break;
                }
                parent = TreeWalker.RawViewWalker.GetParent(parent);
            }

            // 3. 윈도우 핸들을 찾지 못한 경우 현재 활성화된 창 사용
            if (hwnd == IntPtr.Zero)
            {
                hwnd = GetForegroundWindow();
                Console.WriteLine($"[정보] 현재 활성화된 창 핸들 사용: {hwnd}");
            }

            if (hwnd == IntPtr.Zero)
            {
                Console.WriteLine("[오류] 유효한 윈도우 핸들을 찾을 수 없습니다.");
                return;
            }

            // 4. 스레드 연결 및 복사 수행
            uint targetThread = GetWindowThreadProcessId(hwnd, out _);
            uint currentThread = GetCurrentThreadId();

            try
            {
                if (AttachThreadInput(currentThread, targetThread, true))
                {
                    // 윈도우 전경으로
                    SetForegroundWindow(hwnd);
                    Thread.Sleep(100);

                    // 클립보드 초기화
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        try
                        {
                            Clipboard.Clear();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Clipboard.Clear() 오류: {ex.Message}");
                        }
                    });

                    // Ctrl+C 전송
                    SendKeys.SendWait("^c");
                    Thread.Sleep(300);
                }
            }
            finally
            {
                AttachThreadInput(currentThread, targetThread, false);
            }
        }

        private string GetTextFromClipboard()
        {
            // 여러 번 재시도하면서 Clipboard.ContainsText/Clipboard.GetText 사용
            for (int i = 0; i < MAX_RETRIES; i++)
            {
                string text = string.Empty;
                bool got = Application.Current.Dispatcher.Invoke(() =>
                {
                    try
                    {
                        if (Clipboard.ContainsText())
                        {
                            text = Clipboard.GetText();
                            return true;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Clipboard.GetText() 시도 {i + 1} 실패: {ex.Message}");
                    }
                    return false;
                });

                if (got && !string.IsNullOrEmpty(text))
                    return text;

                Thread.Sleep(RETRY_DELAY_MS);
            }

            return string.Empty;
        }

        private void CheckThreadAndPermissions()
        {
            Console.WriteLine($"스레드 ID: {Thread.CurrentThread.ManagedThreadId}");
            Console.WriteLine($"STA 여부: {Thread.CurrentThread.GetApartmentState() == ApartmentState.STA}");
            var id = WindowsIdentity.GetCurrent();
            Console.WriteLine($"사용자: {id.Name}");
            Console.WriteLine($"관리자 권한: {new WindowsPrincipal(id).IsInRole(WindowsBuiltInRole.Administrator)}");
        }

        public override (string SelectedText, Dictionary<string, object> StyleAttributes, string LineNumber)
            GetSelectedTextWithStyle(bool readAllContent = false)
        {
            var style = new Dictionary<string, object>();

            try
            {
                Console.WriteLine("GetSelectedTextWithStyle 시작");
                CheckThreadAndPermissions();

                // 1) UI 스레드에서 복사 동작
                Application.Current.Dispatcher.Invoke(() =>
                {
                    Console.WriteLine("클립보드 복사 시도");
                    CopyToClipboard();
                });

                // 2) 복사된 텍스트 읽기
                Console.WriteLine("클립보드에서 텍스트 읽기");
                string text = GetTextFromClipboard();
                Console.WriteLine($"가져온 텍스트 길이: {text.Length}");

                if (string.IsNullOrEmpty(text))
                    Console.WriteLine("경고: 텍스트가 비어 있습니다.");

                return (text, style, string.Empty);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"오류: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                return (string.Empty, style, string.Empty);
            }
        }

        public override (ulong? FileId, uint? VolumeId, string FileType, string FileName, string FilePath) GetFileInfo()
        {
            return (null, null, "Text", string.Empty, string.Empty);
        }
    }
}
