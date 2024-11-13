using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using Newtonsoft.Json;
using System.IO;
using System.Text;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Diagnostics;

namespace windowsresizer
{
    public class Config
    {
        public int widthIncrement { get; set; } = 10;
        public int heightIncrement { get; set; } = 10;
    }

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private int widthIncrement = 10;
        private int heightIncrement = 10;
        private const string CONFIG_FILE = "config.json";
        private IntPtr mouseHookID = IntPtr.Zero;
        private const int WH_MOUSE_LL = 14;
        private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);
        private readonly LowLevelMouseProc mouseHookProc;

        [StructLayout(LayoutKind.Sequential)]
        private struct MSLLHOOKSTRUCT
        {
            public POINT pt;
            public uint mouseData;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int x;
            public int y;
        }

        public MainWindow()
        {
            InitializeComponent();
            mouseHookProc = MouseHookCallback;
            try
            {
                LoadConfig();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载配置文件失败: {ex.Message}\n将使用默认配置。", "配置加载错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                CreateDefaultConfig();
            }

            mouseHookID = SetHook(mouseHookProc);
            if (mouseHookID == IntPtr.Zero)
            {
                MessageBox.Show("无法设置鼠标钩子", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                Application.Current.Shutdown();
            }

            this.WindowState = WindowState.Minimized;
            this.ShowInTaskbar = false;
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        private IntPtr SetHook(LowLevelMouseProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_MOUSE_LL, proc,
                    GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            const int WM_MOUSEWHEEL = 0x020A;
            if (nCode >= 0 && wParam == (IntPtr)WM_MOUSEWHEEL)
            {
                MSLLHOOKSTRUCT hookStruct = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));
                int delta = (short)((hookStruct.mouseData >> 16) & 0xFFFF);
                bool ctrlPressed = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
                bool altPressed = (GetKeyState(VK_MENU) & 0x8000) != 0;

                if (ctrlPressed || altPressed)
                {
                    IntPtr foregroundWindow = GetForegroundWindow();
                    if (foregroundWindow != IntPtr.Zero && !IsZoomed(foregroundWindow))
                    {
                        RECT rect;
                        GetWindowRect(foregroundWindow, out rect);

                        if (ctrlPressed)
                        {
                            int newWidth = rect.Right - rect.Left + (delta > 0 ? widthIncrement : -widthIncrement);
                            SetWindowPos(foregroundWindow, IntPtr.Zero, rect.Left, rect.Top, newWidth, rect.Bottom - rect.Top, SWP_NOZORDER | SWP_NOMOVE);
                        }
                        else if (altPressed)
                        {
                            int newHeight = rect.Bottom - rect.Top + (delta > 0 ? heightIncrement : -heightIncrement);
                            SetWindowPos(foregroundWindow, IntPtr.Zero, rect.Left, rect.Top, rect.Right - rect.Left, newHeight, SWP_NOZORDER | SWP_NOMOVE);
                        }
                        return (IntPtr)1; // 处理消息
                    }
                }
            }
            return CallNextHookEx(mouseHookID, nCode, wParam, lParam);
        }

        protected override void OnClosed(EventArgs e)
        {
            if (mouseHookID != IntPtr.Zero)
            {
                UnhookWindowsHookEx(mouseHookID);
            }
            base.OnClosed(e);
        }

        private void LoadConfig()
        {
            if (!File.Exists(CONFIG_FILE))
            {
                CreateDefaultConfig();
            }

            var configText = File.ReadAllText(CONFIG_FILE);
            var config = JsonConvert.DeserializeObject<Config>(configText);
            if (config != null)
            {
                widthIncrement = config.widthIncrement;
                heightIncrement = config.heightIncrement;
            }
        }

        private void CreateDefaultConfig()
        {
            var config = new Config();
            var jsonString = JsonConvert.SerializeObject(config, Formatting.Indented);
            File.WriteAllText(CONFIG_FILE, jsonString);
        }

        [DllImport("user32.dll")]
        private static extern short GetKeyState(int nVirtKey);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsZoomed(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        private const int VK_CONTROL = 0x11;
        private const int VK_MENU = 0x12;
        private const uint SWP_NOZORDER = 0x0004;
        private const uint SWP_NOMOVE = 0x0002;

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
    }
}
