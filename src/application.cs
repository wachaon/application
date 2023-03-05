
using System;
using System.Text;
using System.Windows;
using System.Threading;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace wes {
    public class Application {

        [STAThread]
        public static void Main(params string[] args) {
            string method = args[0];

            if (method.StartsWith("window.")) {
                IntPtr hWnd = GetForegroundWindow();
                int SW_SHOWNORMAL = 1;
                int SW_SHOWMINIMIZED = 2;
                int SW_MAXIMIZE = 3;

                if (method == "window.move") {
                    int left = Int32.Parse(args[1]);
                    int top = Int32.Parse(args[2]);
                    int width = Int32.Parse(args[3]);
                    int height = Int32.Parse(args[4]);
                    int right = width - left;
                    int bottom = height - top;
                    bool repaint = true;
                    MoveWindow(hWnd, left, top, width, height, repaint);
                }

                if (method == "window.get") {
                    RECT rect;
                    IntPtr hwnd = GetForegroundWindow();
                    bool flag = GetWindowRect(hwnd, out rect);
                    const int nChars = 256;
                    string title = "";
                    System.Text.StringBuilder Buff = new System.Text.StringBuilder(nChars);
                    if (GetWindowText(hwnd, Buff, nChars) > 0) {
                        title = Buff.ToString();
                    }
                    Console.WriteLine(
                        "{{\"hwnd\": {0}, \"left\": {1}, \"top\": {2}, \"width\": {3}, \"height\": {4}, \"title\": \"{5}\"}}",
                        hwnd,
                        rect.left,
                        rect.top,
                        rect.right - rect.left,
                        rect.bottom - rect.top,
                        title
                    );
                }

                if (method == "window.normal") {
                    IntPtr hwnd = GetForegroundWindow();
                    ShowWindowAsync(hWnd, SW_SHOWNORMAL);
                }

                if (method == "window.min") {
                    IntPtr hwnd = GetForegroundWindow();
                    ShowWindowAsync(hWnd, SW_SHOWMINIMIZED);
                }

                if (method == "window.max") {
                    IntPtr hwnd = GetForegroundWindow();
                    ShowWindowAsync(hWnd, SW_MAXIMIZE);
                }

                if (method == "window.activate.window.title") {
                    IntPtr hwnd = FindWindow(null, args[1]);
                    if (hwnd != IntPtr.Zero) {
                        SetForegroundWindow(hwnd);
                        ShowWindowAsync(hwnd, SW_SHOWNORMAL);
                        SetActiveWindow(hwnd);
                    }
                }

                if (method == "window.activate.window.handle") {
                    IntPtr hwnd = new IntPtr(Int32.Parse(args[1]));
                    if (hwnd != IntPtr.Zero) {
                        SetForegroundWindow(hwnd);
                        ShowWindowAsync(hwnd, SW_SHOWNORMAL);
                        SetActiveWindow(hwnd);
                    }
                }
            }

            if (method.StartsWith("keyboard.")) {
                uint KEYEVENTF_EXTENDEDKEY = 0x1;
                uint KEYEVENTF_KEYUP = 0x2;
                int key_code = Int32.Parse(args[1]);
                byte input = (byte)key_code;

                if (method == "keyboard.send") {
                    keybd_event(input, 0, KEYEVENTF_EXTENDEDKEY, 0);
                }

                if (method == "keyboard.press") {
                    keybd_event(input, 0, KEYEVENTF_EXTENDEDKEY, 0);
                    if (args.Length > 2) {
                        int timer = Int32.Parse(args[2]);
                        Thread.Sleep(timer);
                    }
                }

                if (method == "keyboard.send" || method == "keyboard.release") {
                    keybd_event(input, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0);
                }
            }

            if (method.StartsWith("mouse.")) {
                int posX = args.Length > 1 ? Int32.Parse(args[1]) : 0;
                int posY = args.Length > 2 ? Int32.Parse(args[2]) : 0;

                int MOUSEEVENTF_LEFTDOWN = 0x0002;
                int MOUSEEVENTF_LEFTUP = 0x0004;
                int MOUSEEVENTF_RIGHTDOWN = 0x0008;
                int MOUSEEVENTF_RIGHTUP = 0x0010;
                int MOUSEEVENTF_MIDDLEDOWN = 0x0020;
                int MOUSEEVENTF_MIDDLEUP = 0x0040;
                int MOUSEEVENTF_WHEEL = 0x0800;

                if (method == "mouse.pos") { SetCursorPos(posX, posY); }

                if (method == "mouse.click" || method == "mouse.leftDown") { mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0); }
                if (method == "mouse.click" || method == "mouse.leftUp"  ) { mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0); }

                if (method == "mouse.rightClick" || method == "mouse.rightDown") { mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0); }
                if (method == "mouse.rightClick" || method == "mouse.rightUp"  ) { mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0); }

                if (method == "mouse.middleClick" || method == "mouse.middleDown") { mouse_event(MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0); }
                if (method == "mouse.middleClick" || method == "mouse.middleUp"  ) { mouse_event(MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0); }

                if (method == "mouse.scroll") { mouse_event(MOUSEEVENTF_WHEEL,0,0,posX,0); }
            }

            if (method.StartsWith("clipboard.")) {
                string format = args[1];
                if (method == "clipboard.setData") {
                    if (format == "Text") {
                        string data = args[2];
                        Clipboard.SetDataObject(data, true);
                    }
                }
                if (method == "clipboard.getData") {
                    if (format == "Text") {
                        IDataObject data = Clipboard.GetDataObject();
                        if(data.GetDataPresent(DataFormats.Text)) {
                            Console.WriteLine(data.GetData(DataFormats.Text));
                        }
                    }
                }
            }
        }

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern int MoveWindow(IntPtr hwnd, int x, int y, int nWidth,int nHeight, bool bRepaint);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool ShowWindowAsync(IntPtr hwnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern IntPtr SetActiveWindow(IntPtr hwnd);

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpWindowText, int nMaxCount);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }

        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, uint dwExtraInfo);

        [DllImport("user32.dll")]
        public static extern void SetCursorPos(int X, int Y);

        [DllImport("user32.dll")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

    }
}
