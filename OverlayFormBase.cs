using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace PPTTool
{
    /// <summary>
    /// 覆盖窗口基类
    /// - 分层窗口（半透明）
    /// - 按钮区域可点击，非按钮区域可拖拽移动（HTCAPTION）
    /// - 点击按钮后自动将焦点归还给 PowerPoint
    /// - 始终置顶、无边框、不显示在任务栏
    /// </summary>
    public class OverlayFormBase : Form
    {
        protected PowerPointService Service { get; }

        /// <summary>可点击的热区列表（按钮区域）</summary>
        private readonly List<Rectangle> _hitAreas = new List<Rectangle>();

        /// <summary>PPT 放映窗口句柄，用于归还焦点</summary>
        private IntPtr _pptHwnd = IntPtr.Zero;

        public OverlayFormBase(PowerPointService service)
        {
            Service = service ?? throw new ArgumentNullException(nameof(service));

            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
            BackColor = Color.Black;
            TransparencyKey = Color.Black;
            MinimizeBox = false;
            MaximizeBox = false;
            ControlBox = false;

            SetStyle(ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.UserPaint |
                     ControlStyles.DoubleBuffer |
                     ControlStyles.ResizeRedraw, true);
        }

        /// <summary>获取 PPT 放映窗口句柄</summary>
        protected void CapturePptHwnd()
        {
            try
            {
                _pptHwnd = Service.GetSlideShowHwnd();
            }
            catch
            {
                _pptHwnd = IntPtr.Zero;
            }
        }

        /// <summary>将焦点归还给 PowerPoint</summary>
        protected void FocusBackToPpt()
        {
            if (_pptHwnd != IntPtr.Zero)
            {
                NativeMethods.SetFocus(_pptHwnd);
            }
        }

        /// <summary>注册一个可点击的热区（按钮区域）</summary>
        protected void AddHitArea(Rectangle area)
        {
            _hitAreas.Add(area);
        }

        /// <summary>清除所有热区</summary>
        protected void ClearHitAreas()
        {
            _hitAreas.Clear();
        }

        /// <summary>判断某个客户端坐标是否在按钮上</summary>
        protected bool IsOverButton(Point clientPt)
        {
            foreach (var area in _hitAreas)
            {
                if (area.Contains(clientPt))
                    return true;
            }
            return false;
        }

        // ─── 点击穿透 + 拖拽移动 核心逻辑 ───
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == NativeMethods.WM_NCHITTEST)
            {
                int x = (short)(m.LParam.ToInt32() & 0xFFFF);
                int y = (short)((m.LParam.ToInt32() >> 16) & 0xFFFF);
                Point pt = PointToClient(new Point(x, y));

                if (IsOverButton(pt))
                {
                    // 按钮区域 → 可点击
                    m.Result = (IntPtr)NativeMethods.HTCLIENT;
                }
                else
                {
                    // 非按钮区域 → 模拟标题栏拖拽，实现窗口移动
                    m.Result = (IntPtr)NativeMethods.HTCAPTION;
                }
                return;
            }

            base.WndProc(ref m);
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            if (!DesignMode)
            {
                NativeMethods.MakeLayeredOverlay(this, 200);
                CapturePptHwnd();
            }
        }

        /// <summary>
        /// 点击后归还焦点给 PPT（子类在按钮点击后调用）
        /// </summary>
        protected void OnButtonClicked()
        {
            FocusBackToPpt();
        }
    }
}
