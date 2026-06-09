using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointApp = Microsoft.Office.Interop.PowerPoint.Application;

namespace PPTTool
{
    /// <summary>
    /// 幻灯片放映跟踪器
    /// 监听 SlideShowBegin/SlideShowEnd 事件，创建/销毁工具栏和翻页栏
    /// </summary>
    public class SlideShowTracker : IDisposable
    {
        private readonly PowerPointApp _pptApp;
        private PowerPointService _service;
        private ToolbarOverlay _toolbar;
        private NavOverlay _leftNav;   // 左下角翻页栏（含上+下）
        private NavOverlay _rightNav;  // 右下角翻页栏（含上+下）
        private bool _disposed;
        private Timer _focusTimer;

        public SlideShowTracker(PowerPointApp pptApp)
        {
            _pptApp = pptApp ?? throw new ArgumentNullException(nameof(pptApp));
            _pptApp.SlideShowBegin += OnSlideShowBegin;
            _pptApp.SlideShowEnd += OnSlideShowEnd;
        }

        private void OnSlideShowBegin(SlideShowWindow wn)
        {
            if (_disposed) return;

            try
            {
                _service = new PowerPointService(wn);
                _service.SetPointerArrow();

                // 工具栏（底部居中）
                _toolbar = new ToolbarOverlay(_service);
                _toolbar.Show();

                // 翻页栏 × 2（左下角 + 右下角，各含上一页+下一页）
                _leftNav = new NavOverlay(_service, NavOverlay.NavPosition.Left);
                _leftNav.Show();

                _rightNav = new NavOverlay(_service, NavOverlay.NavPosition.Right);
                _rightNav.Show();

                // 置顶
                NativeMethods.KeepTopMost(_toolbar);
                NativeMethods.KeepTopMost(_leftNav);
                NativeMethods.KeepTopMost(_rightNav);

                StartFocusTimer();

                System.Diagnostics.Debug.WriteLine("[PPTTool] 放映工具栏已启动");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PPTTool] 启动失败: {ex.Message}");
            }
        }

        private void OnSlideShowEnd(Presentation pres)
        {
            if (_disposed) return;

            try
            {
                StopFocusTimer();

                SafeClose(ref _toolbar);
                SafeClose(ref _leftNav);
                SafeClose(ref _rightNav);

                if (_service != null)
                {
                    _service.Dispose();
                    _service = null;
                }

                System.Diagnostics.Debug.WriteLine("[PPTTool] 放映工具栏已关闭");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PPTTool] 关闭失败: {ex.Message}");
            }
        }

        private void SafeClose<T>(ref T form) where T : Form
        {
            if (form != null)
            {
                form.Close();
                form.Dispose();
                form = null;
            }
        }

        private void StartFocusTimer()
        {
            _focusTimer = new Timer { Interval = 500 };
            _focusTimer.Tick += (s, e) =>
            {
                try
                {
                    if (_service != null && _service.IsInShow)
                    {
                        IntPtr pptHwnd = _service.GetSlideShowHwnd();
                        if (pptHwnd != IntPtr.Zero)
                            NativeMethods.SetFocus(pptHwnd);
                    }
                }
                catch { }
            };
            _focusTimer.Start();
        }

        private void StopFocusTimer()
        {
            if (_focusTimer != null)
            {
                _focusTimer.Stop();
                _focusTimer.Dispose();
                _focusTimer = null;
            }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _disposed = true;

                if (_pptApp != null)
                {
                    try
                    {
                        _pptApp.SlideShowBegin -= OnSlideShowBegin;
                        _pptApp.SlideShowEnd -= OnSlideShowEnd;
                    }
                    catch { }
                }

                OnSlideShowEnd(null);
            }
        }
    }
}
