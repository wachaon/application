using System;
using System.Runtime.InteropServices;

public class keyboard {
    public static void Main (params string[] args) {
        string method = args[0];
        int key_code = Int32.Parse(args[1]);
        byte input = (byte)key_code;

        uint KEYEVENTF_EXTENDEDKEY = 0x1;
        uint KEYEVENTF_KEYUP = 0x2;

        if (method == "send" || method == "press") { keybd_event(input, 0, KEYEVENTF_EXTENDEDKEY, 0); }
        if (method == "send" || method == "release") { keybd_event(input, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0); }
    }

    [DllImport("user32.dll")]
    private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, uint dwExtraInfo);
}
