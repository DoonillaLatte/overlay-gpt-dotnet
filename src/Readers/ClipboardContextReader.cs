using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using System.Windows.Automation;
using System.Runtime.InteropServices;
using System.Threading;

namespace overlay_gpt
{
    public class ClipboardContextReader : BaseContextReader
    {
        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool CloseClipboard();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool EmptyClipboard();

        private const int KEYEVENTF_EXTENDEDKEY = 0x0001;
        private const int KEYEVENTF_KEYUP = 0x0002;
        private const byte VK_CONTROL = 0x11;
        private const byte VK_A = 0x41;
        private const byte VK_C = 0x43;
        private const int MAX_RETRIES = 3;
        private const int RETRY_DELAY = 100;

        private void SimulateKeyPress(byte key)
        {
            try
            {
                LogWindow.Instance.Log($"키 입력 시뮬레이션 시작: {key}");
                keybd_event(key, 0, KEYEVENTF_EXTENDEDKEY, 0);
                Thread.Sleep(100);
                keybd_event(key, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0);
                Thread.Sleep(100);
                LogWindow.Instance.Log($"키 입력 시뮬레이션 완료: {key}");
            }
            catch (System.Exception ex)
            {
                LogWindow.Instance.Log($"키 입력 시뮬레이션 중 오류: {ex.Message}");
                throw;
            }
        }

        private bool TryOpenClipboard()
        {
            for (int i = 0; i < MAX_RETRIES; i++)
            {
                if (OpenClipboard(IntPtr.Zero))
                {
                    return true;
                }
                LogWindow.Instance.Log($"클립보드 열기 시도 {i + 1}/{MAX_RETRIES} 실패");
                Thread.Sleep(RETRY_DELAY);
            }
            return false;
        }

        private void CopyToClipboard()
        {
            try
            {
                LogWindow.Instance.Log("클립보드 복사 프로세스 시작");
                
                var focusedElement = AutomationElement.FocusedElement;
                if (focusedElement == null)
                {
                    LogWindow.Instance.Log("경고: 포커스된 요소가 없습니다");
                    return;
                }

                // 클립보드 초기화
                if (TryOpenClipboard())
                {
                    EmptyClipboard();
                    CloseClipboard();
                }
                
                LogWindow.Instance.Log("Ctrl 키 누르기 시작");
                keybd_event(VK_CONTROL, 0, KEYEVENTF_EXTENDEDKEY, 0);
                Thread.Sleep(100);
                LogWindow.Instance.Log("Ctrl 키 누르기 완료");

                LogWindow.Instance.Log("Ctrl+C 실행 시작");
                SimulateKeyPress(VK_C);
                LogWindow.Instance.Log("Ctrl+C 실행 완료");

                LogWindow.Instance.Log("Ctrl 키 떼기 시작");
                keybd_event(VK_CONTROL, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0);
                Thread.Sleep(200);
                LogWindow.Instance.Log("Ctrl 키 떼기 완료");
                
                Thread.Sleep(300);
                LogWindow.Instance.Log("클립보드 복사 프로세스 완료");
            }
            catch (System.Exception ex)
            {
                LogWindow.Instance.Log($"복사 중 오류 발생: {ex.Message}");
                LogWindow.Instance.Log($"스택 트레이스: {ex.StackTrace}");
                throw;
            }
        }

        private string GetTextFromClipboard()
        {
            for (int i = 0; i < MAX_RETRIES; i++)
            {
                try
                {
                    if (TryOpenClipboard())
                    {
                        try
                        {
                            string text = Clipboard.GetText();
                            return text ?? string.Empty;
                        }
                        finally
                        {
                            CloseClipboard();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    LogWindow.Instance.Log($"클립보드 읽기 시도 {i + 1}/{MAX_RETRIES} 실패: {ex.Message}");
                }
                Thread.Sleep(RETRY_DELAY);
            }
            return string.Empty;
        }

        public override (string SelectedText, Dictionary<string, object> StyleAttributes, string LineNumber) GetSelectedTextWithStyle(bool readAllContent = false)
        {
            var styleAttributes = new Dictionary<string, object>();
            
            try
            {
                LogWindow.Instance.Log($"GetSelectedTextWithStyle 호출됨 (readAllContent: {readAllContent})");
                
                LogWindow.Instance.Log("클립보드 복사 시작");
                CopyToClipboard();
                
                LogWindow.Instance.Log("클립보드에서 텍스트 가져오기 시작");
                string text = GetTextFromClipboard();
                LogWindow.Instance.Log($"클립보드에서 텍스트 가져오기 완료 - 길이: {text?.Length ?? 0}");
                
                if (string.IsNullOrEmpty(text))
                {
                    LogWindow.Instance.Log("경고: 가져온 텍스트가 비어있습니다");
                }
                
                return (text, styleAttributes, string.Empty);
            }
            catch (System.Exception ex)
            {
                LogWindow.Instance.Log($"텍스트 가져오기 중 오류 발생: {ex.Message}");
                LogWindow.Instance.Log($"스택 트레이스: {ex.StackTrace}");
                return (string.Empty, styleAttributes, string.Empty);
            }
        }
    }
} 