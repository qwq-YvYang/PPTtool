using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PPTTool
{
    internal static class NativeMethods
    {
        // ─── 窗口样式常量 ───
        internal const int GWL_EXSTYLE = -20;
        internal const int WS_EX_LAYERED = 0x80000;
        internal const int WS_EX_TRANSPARENT = 0x20;
        internal const int WS_EX_TOOLWINDOW = 0x80;
        internal const int WS_EX_TOPMOST = 0x8;
        internal const int WS_EX_NOACTIVATE = 0x08000000;

        // ─── 分层窗口常量 ───
        internal const uint LWA_ALPHA = 0x2;

        // ─── WM_NCHITTEST 返回值 ───
        internal const int HTCLIENT = 1;
        internal const int HTCAPTION = 2;
        internal const int HTTRANSPARENT = -1;

        // ─── 窗口消息 ───
        internal const int WM_NCHITTEST = 0x84;
        internal const int WM_MOUSEMOVE = 0x0200;
        internal const int WM_LBUTTONDOWN = 0x0201;
        internal const int WM_LBUTTONUP = 0x0202;

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool SetLayeredWindowAttributes(
            IntPtr hwnd, uint crKey, byte bAlpha, uint dwFlags);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool SetWindowPos(
            IntPtr hWnd, IntPtr hWndInsertAfter,
            int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr SetFocus(IntPtr hWnd);

        internal static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        internal const uint SWP_NOSIZE = 0x0001;
        internal const uint SWP_NOMOVE = 0x0002;
        internal const uint SWP_NOACTIVATE = 0x0010;
        internal const uint SWP_SHOWWINDOW = 0x0040;

        internal static void MakeLayeredOverlay(Form form, byte opacity = 200)
        {
            IntPtr hwnd = form.Handle;
            int style = GetWindowLong(hwnd, GWL_EXSTYLE);
            style |= WS_EX_LAYERED | WS_EX_TOOLWINDOW | WS_EX_TOPMOST | WS_EX_NOACTIVATE;
            SetWindowLong(hwnd, GWL_EXSTYLE, style);
            SetLayeredWindowAttributes(hwnd, 0, opacity, LWA_ALPHA);
        }

        internal static void KeepTopMost(Form form)
        {
            SetWindowPos(form.Handle, HWND_TOPMOST, 0, 0, 0, 0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW);
        }
    }
}
