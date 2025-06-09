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
        // Win32 API: 포커스 강제용
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
            var focused = AutomationElement.FocusedElement;
            if (focused == null)
            {
                Console.WriteLine("[오류] 포커스된 요소가 없습니다.");
                return;
            }
            Console.WriteLine($"[정보] 포커스된 요소: {focused.Current.Name} (ControlType: {focused.Current.ControlType.ProgrammaticName})");

            // TextPattern 시도
            Console.WriteLine("\n[시도] TextPattern으로 텍스트 가져오기");
            try
            {
                var textPattern = focused.GetCurrentPattern(TextPattern.Pattern) as TextPattern;
                if (textPattern != null)
                {
                    Console.WriteLine("[성공] TextPattern 패턴 획득");
                    var text = textPattern.DocumentRange.GetText(-1);
                    Console.WriteLine($"[정보] TextPattern으로 가져온 텍스트 길이: {text?.Length ?? 0}");
                    
                    if (!string.IsNullOrEmpty(text))
                    {
                        Console.WriteLine("[시도] TextPattern 텍스트를 클립보드에 설정");
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            try
                            {
                                Clipboard.SetText(text);
                                Console.WriteLine("[성공] TextPattern 텍스트 클립보드 설정 완료");
                                return;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"[오류] TextPattern 클립보드 설정 실패: {ex.Message}");
                                Console.WriteLine($"[상세] {ex.StackTrace}");
                            }
                        });
                    }
                    else
                    {
                        Console.WriteLine("[알림] TextPattern으로 가져온 텍스트가 비어있음");
                    }
                }
                else
                {
                    Console.WriteLine("[알림] TextPattern을 지원하지 않는 요소");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[오류] TextPattern 시도 실패: {ex.Message}");
                Console.WriteLine($"[상세] {ex.StackTrace}");
            }

            // ValuePattern 시도
            Console.WriteLine("\n[시도] ValuePattern으로 텍스트 가져오기");
            try
            {
                var valuePattern = focused.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                if (valuePattern != null)
                {
                    Console.WriteLine("[성공] ValuePattern 패턴 획득");
                    var value = valuePattern.Current.Value;
                    Console.WriteLine($"[정보] ValuePattern으로 가져온 텍스트 길이: {value?.Length ?? 0}");
                    
                    if (!string.IsNullOrEmpty(value))
                    {
                        Console.WriteLine("[시도] ValuePattern 텍스트를 클립보드에 설정");
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            try
                            {
                                Clipboard.SetText(value);
                                Console.WriteLine("[성공] ValuePattern 텍스트 클립보드 설정 완료");
                                return;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"[오류] ValuePattern 클립보드 설정 실패: {ex.Message}");
                                Console.WriteLine($"[상세] {ex.StackTrace}");
                            }
                        });
                    }
                    else
                    {
                        Console.WriteLine("[알림] ValuePattern으로 가져온 텍스트가 비어있음");
                    }
                }
                else
                {
                    Console.WriteLine("[알림] ValuePattern을 지원하지 않는 요소");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[오류] ValuePattern 시도 실패: {ex.Message}");
                Console.WriteLine($"[상세] {ex.StackTrace}");
            }

            // 기존 클립보드 복사 방식 시도
            Console.WriteLine("\n[시도] Ctrl+C 방식으로 텍스트 가져오기");
            IntPtr hwnd = new IntPtr(focused.Current.NativeWindowHandle);
            if (hwnd == IntPtr.Zero)
            {
                Console.WriteLine("[오류] 윈도우 핸들을 가져올 수 없습니다.");
                return;
            }
            Console.WriteLine($"[정보] 윈도우 핸들: {hwnd}");

            // 스레드 연결
            uint targetThread = GetWindowThreadProcessId(hwnd, out _);
            uint currentThread = GetCurrentThreadId();
            
            try
            {
                if (AttachThreadInput(currentThread, targetThread, true))
                {
                    // 윈도우 전경으로
                    SetForegroundWindow(hwnd);
                    Thread.Sleep(100);
                    focused.SetFocus();
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
                // 스레드 연결 해제
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
