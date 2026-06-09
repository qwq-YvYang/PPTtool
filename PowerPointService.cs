using System.Runtime.InteropServices;
using System;
using System.Drawing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PPTTool
{
    /// <summary>
    /// 封装所有 PowerPoint 原生 API 调用
    /// 所有功能均通过 COM 接口调用，无模拟点击
    /// </summary>
    public class PowerPointService : IDisposable
    {
        private readonly SlideShowWindow _window;
        private bool _disposed;

        /// <summary>放映视图对象</summary>
        public SlideShowView View => _window.View;

        /// <summary>当前是否在放映中</summary>
        public bool IsInShow => _window != null;

        public PowerPointService(SlideShowWindow window)
        {
            _window = window ?? throw new ArgumentNullException(nameof(window));
        }

        // ─── 指针模式 ───

        /// <summary>切换为箭头指针</summary>
        public void SetPointerArrow()
        {
            CheckDisposed();
            _window.View.PointerType = PpSlideShowPointerType.ppSlideShowPointerArrow;
        }

        /// <summary>切换为画笔</summary>
        public void SetPointerPen()
        {
            CheckDisposed();
            _window.View.PointerType = PpSlideShowPointerType.ppSlideShowPointerPen;
        }

        /// <summary>切换为橡皮擦</summary>
        public void SetPointerEraser()
        {
            CheckDisposed();
            _window.View.PointerType = PpSlideShowPointerType.ppSlideShowPointerEraser;
        }

        /// <summary>判断当前是否为画笔模式</summary>
        public bool IsPenMode =>
            _window.View.PointerType == PpSlideShowPointerType.ppSlideShowPointerPen;

        /// <summary>判断当前是否为橡皮擦模式</summary>
        public bool IsEraserMode =>
            _window.View.PointerType == PpSlideShowPointerType.ppSlideShowPointerEraser;

        // ─── 画笔颜色 ───

        /// <summary>设置画笔颜色</summary>
        public void SetPenColor(Color color)
        {
            CheckDisposed();
            // RGB 需转成 BGR (PowerPoint 使用 BGR 格式)
            _window.View.PointerColor.RGB = (int)((color.B << 16) | (color.G << 8) | color.R);
        }

        // ─── 笔迹管理 ───

        /// <summary>清除当前页所有笔迹</summary>
        public void EraseDrawing()
        {
            CheckDisposed();
            _window.View.EraseDrawing();
        }

        // ─── 翻页 ───

        /// <summary>下一页</summary>
        public void NextSlide()
        {
            CheckDisposed();
            _window.View.Next();
        }

        /// <summary>上一页</summary>
        public void PreviousSlide()
        {
            CheckDisposed();
            _window.View.Previous();
        }

        // ─── 退出放映 ───

        /// <summary>退出幻灯片放映</summary>
        public void ExitShow()
        {
            CheckDisposed();
            _window.View.Exit();
        }

        // ─── 获取放映窗口句柄和位置 ───

        /// <summary>获取放映窗口句柄</summary>
        public IntPtr GetSlideShowHwnd()
        {
            CheckDisposed();
            return (IntPtr)_window.HWND;
        }

        /// <summary>获取放映窗口的位置和大小</summary>
        public Rectangle GetSlideShowBounds()
        {
            CheckDisposed();
            var left = (int)_window.Left;
            var top = (int)_window.Top;
            var width = (int)_window.Width;
            var height = (int)_window.Height;
            return new Rectangle(left, top, width, height);
        }


        private void CheckDisposed()
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(PowerPointService));
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                // 释放 COM 引用
                if (_window != null)
                    Marshal.ReleaseComObject(_window);
                _disposed = true;
            }
        }
    }
}
