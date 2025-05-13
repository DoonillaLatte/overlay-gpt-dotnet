using System;
using System.Runtime.InteropServices;
using System.Windows.Interop;

namespace overlay_gpt 
{
    public class HotkeyManager
    {
    
        private const int HOTKEY_ID = 9000;
        private const uint MOD_CONTROL = 0x0002;
        private const uint MOD_ALT = 0x0001;
        private const uint VK_K = 0x4B;
    
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        public static void RegisterHotKey(WindowInteropHelper windowHelper, Action onHotKeyPressed)
        {
            var handle = windowHelper.Handle;
            RegisterHotKey(handle, HOTKEY_ID, MOD_CONTROL | MOD_ALT, VK_K);

            HwndSource source = HwndSource.FromHwnd(handle);
            source.AddHook((IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled) =>
            {
                const int WM_HOTKEY = 0x0312;
                if (msg == WM_HOTKEY && wParam.ToInt32() == HOTKEY_ID)
                {
                    onHotKeyPressed?.Invoke();
                    handled = true;
                }
                return IntPtr.Zero;
            });
        }
    }
}
