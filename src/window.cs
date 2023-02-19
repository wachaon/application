using System;
using System.Windows;
using System.Runtime.InteropServices;

public class window {
    public static int Main (params string[] args) {
        string method = args[0];
        IntPtr hWnd = GetForegroundWindow();

        int SW_SHOWNORMAL = 1;
        int SW_SHOWMINIMIZED = 2;
        int SW_MAXIMIZE = 3;

        if (method == "move") {
            int left = Int32.Parse(args[1]);
            int top = Int32.Parse(args[2]);
            int width = Int32.Parse(args[3]);
            int height = Int32.Parse(args[4]);
            int right = width - left;
            int bottom = height - top;
            bool repaint = true;
            MoveWindow(hWnd, left, top, width, height, repaint);
        }
        if (method == "get") {
            RECT rect;
            bool flag = GetWindowRect(hWnd, out rect);

            if (args.Length == 1) {
                Console.WriteLine(
                    "hWnd: {0} left: {1} top: {2} width: {3} height: {4}",
                    hWnd,
                    rect.left,
                    rect.top,
                    rect.right - rect.left,
                    rect.bottom - rect.top
                );
            } else {
                return 0;
            }
        }
        if (method == "hwnd") {
            if (args.Length == 1) {
                Console.WriteLine(hWnd);
            } else {
                return Int32.Parse(hWnd.ToString());
            }
        }
        if (method == "windowLeft") {
            RECT rect;
            bool flag = GetWindowRect(hWnd, out rect);

            if (args.Length == 1) {
                Console.WriteLine("left {0}", rect.left);
            } else {
                return Int32.Parse(rect.left.ToString());
            }
        }
        if (method == "windowTop") {
            RECT rect;
            bool flag = GetWindowRect(hWnd, out rect);

            if (args.Length == 1) {
                Console.WriteLine("top: {0}", rect.top);
            } else {
                return Int32.Parse(rect.top.ToString());
            }
        }
        if (method == "windowWidth") {
            RECT rect;
            bool flag = GetWindowRect(hWnd, out rect);

            if (args.Length == 1) {
                Console.WriteLine("whidth: {0}", rect.right - rect.left);
            } else {
                return Int32.Parse((rect.right - rect.left).ToString());
            }
        }
        if (method == "windowHeight") {
            RECT rect;
            bool flag = GetWindowRect(hWnd, out rect);

            if (args.Length == 1) {
                Console.WriteLine("height: {0}",rect.bottom - rect.top);
            } else {
                return Int32.Parse((rect.bottom - rect.top).ToString());
            }
        }
        if (method == "active") {
            IntPtr hwnd = new IntPtr(Int32.Parse(args[1]));
            ShowWindowAsync(hwnd, SW_SHOWNORMAL);
            SetActiveWindow(hwnd);
        }

        if (method == "normal") {
            ShowWindowAsync(hWnd, SW_SHOWNORMAL);
        }
        if (method == "min") {
            ShowWindowAsync(hWnd, SW_SHOWMINIMIZED);
        }
        if (method == "max") {
            ShowWindowAsync(hWnd, SW_MAXIMIZE);
        }
        return 0;
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

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT {
        public int left;
        public int top;
        public int right;
        public int bottom;
    }
}
