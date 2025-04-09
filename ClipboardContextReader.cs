using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using System.Windows.Automation;
using System.Runtime.InteropServices;

namespace overlay_gpt
{
    public class ClipboardContextReader : BaseContextReader
    {
        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);

        private const int KEYEVENTF_EXTENDEDKEY = 0x0001;
        private const int KEYEVENTF_KEYUP = 0x0002;
        private const byte VK_CONTROL = 0x11;
        private const byte VK_A = 0x41;
        private const byte VK_C = 0x43;

        private void SimulateKeyPress(byte key)
        {
            keybd_event(key, 0, KEYEVENTF_EXTENDEDKEY, 0);
            keybd_event(key, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0);
        }

        private void CopyToClipboard()
        {
            try
            {
                // Ctrl 키 누르기
                keybd_event(VK_CONTROL, 0, KEYEVENTF_EXTENDEDKEY, 0);

                // 현재 포커스된 요소 가져오기
                var focusedElement = AutomationElement.FocusedElement;
                if (focusedElement != null)
                {
                    // TextPattern이 있는지 확인
                    if (focusedElement.TryGetCurrentPattern(TextPattern.Pattern, out object pattern))
                    {
                        var textPattern = (TextPattern)pattern;
                        var selection = textPattern.GetSelection();
                        
                        // 선택된 텍스트가 있는지 확인
                        if (selection != null && selection.Length > 0 && selection[0].GetText(-1).Length > 0)
                        {
                            // 선택된 텍스트가 있으면 Ctrl+C
                            SimulateKeyPress(VK_C);
                            LogWindow.Instance.Log("선택된 텍스트 복사 (Ctrl+C)");
                        }
                        else
                        {
                            // 선택된 텍스트가 없으면 Ctrl+A 후 Ctrl+C
                            SimulateKeyPress(VK_A);
                            System.Threading.Thread.Sleep(50); // 약간의 지연
                            SimulateKeyPress(VK_C);
                            LogWindow.Instance.Log("전체 텍스트 복사 (Ctrl+A, Ctrl+C)");
                        }
                    }
                    else
                    {
                        // TextPattern이 없으면 그냥 Ctrl+C 시도
                        SimulateKeyPress(VK_C);
                        LogWindow.Instance.Log("기본 복사 시도 (Ctrl+C)");
                    }
                }

                // Ctrl 키 떼기
                keybd_event(VK_CONTROL, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0);
                
                // 클립보드가 업데이트될 때까지 잠시 대기
                System.Threading.Thread.Sleep(100);
            }
            catch (System.Exception ex)
            {
                LogWindow.Instance.Log($"복사 중 오류 발생: {ex.Message}");
            }
        }

        public override (string SelectedText, Dictionary<string, object> StyleAttributes) GetSelectedTextWithStyle()
        {
            var styleAttributes = new Dictionary<string, object>();
            string selectedText = string.Empty;

            try
            {
                // 먼저 복사 시도
                CopyToClipboard();

                // 클립보드에서 텍스트 읽기
                if (Clipboard.ContainsText())
                {
                    selectedText = Clipboard.GetText();
                    LogWindow.Instance.Log($"클립보드 텍스트: {selectedText} (길이: {selectedText.Length})");
                }
                else
                {
                    LogWindow.Instance.Log("클립보드에 텍스트가 없습니다.");
                }
            }
            catch (System.Exception ex)
            {
                LogWindow.Instance.Log($"클립보드 읽기 오류: {ex.Message}");
            }

            return (selectedText, styleAttributes);
        }
    }
} 