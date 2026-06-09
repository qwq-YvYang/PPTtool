using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace PPTTool
{
    /// <summary>
    /// 底部工具栏（白色半透明圆形边框 + 自定义矢量图标）
    /// </summary>
    public class ToolbarOverlay : OverlayFormBase
    {
        private enum ToolButton { Arrow, Pen, Eraser, Exit }

        private readonly Dictionary<ToolButton, Rectangle> _buttonRects = new Dictionary<ToolButton, Rectangle>();

        // ─── 画笔颜色选择 ───
        private bool _showColorPicker;
        private Color _penColor = Color.Red;
        private readonly Color[] _presetColors = {
            Color.Red, Color.Blue, Color.Green, Color.Yellow,
            Color.Orange, Color.Purple, Color.Black, Color.White
        };
        private Rectangle _colorPanelRect;
        private readonly Dictionary<int, Rectangle> _colorRects = new Dictionary<int, Rectangle>();

        // ─── 橡皮擦子菜单 ───
        private bool _showEraserMenu;
        private Rectangle _eraserMenuRect;

        // ─── 尺寸 ───
        private const int BtnSize = 48;
        private const int Pad = 10;
        private const int EdgePad = 14;   // 边缘留白便于拖动
        private const int BorderWidth = 2;

        // ─── 悬停 ───
        private ToolButton? _hoverButton;

        // ─── UI 颜色配置（后期可自定义） ───
        public Color BorderColor { get; set; } = Color.FromArgb(160, 255, 255, 255); // 白色半透明边框
        public Color BgColor { get; set; } = Color.FromArgb(30, 255, 255, 255);      // 极淡白色背景
        public Color HoverBorderColor { get; set; } = Color.FromArgb(220, 255, 255, 255);
        public Color HoverBgColor { get; set; } = Color.FromArgb(80, 255, 255, 255);
        public Color ActiveBorderColor { get; set; } = Color.FromArgb(220, 100, 180, 255); // 激活态蓝色边框
        public Color IconColor { get; set; } = Color.FromArgb(200, 255, 255, 255);
        public Color HoverIconColor { get; set; } = Color.FromArgb(255, 255, 255, 255);

        public ToolbarOverlay(PowerPointService service) : base(service)
        {
            int totalWidth = EdgePad * 2 + BtnSize * 4 + Pad * 3;
            int totalHeight = EdgePad * 2 + BtnSize;
            Size = new Size(totalWidth, totalHeight);

            var screen = Screen.PrimaryScreen.WorkingArea;
            Location = new Point(
                (screen.Width - Width) / 2,
                screen.Height - Height - 20
            );

            DoubleBuffered = true;
            Paint += ToolbarOverlay_Paint;
            MouseClick += ToolbarOverlay_MouseClick;
            MouseMove += ToolbarOverlay_MouseMove;
            MouseLeave += (s, e) => { _hoverButton = null; Invalidate(); };

            LayoutButtons();
        }

        private void LayoutButtons()
        {
            _buttonRects.Clear();
            ClearHitAreas();

            int x = EdgePad;
            int y = EdgePad;

            _buttonRects[ToolButton.Arrow] = new Rectangle(x, y, BtnSize, BtnSize); x += BtnSize + Pad;
            _buttonRects[ToolButton.Pen] = new Rectangle(x, y, BtnSize, BtnSize); x += BtnSize + Pad;
            _buttonRects[ToolButton.Eraser] = new Rectangle(x, y, BtnSize, BtnSize); x += BtnSize + Pad;
            _buttonRects[ToolButton.Exit] = new Rectangle(x, y, BtnSize, BtnSize);

            foreach (var r in _buttonRects.Values)
                AddHitArea(r);
        }

        private void ToolbarOverlay_Paint(object sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;

            DrawButton(g, ToolButton.Arrow);
            DrawButton(g, ToolButton.Pen);
            DrawButton(g, ToolButton.Eraser);
            DrawButton(g, ToolButton.Exit);

            if (_showColorPicker) DrawColorPicker(g);
            if (_showEraserMenu) DrawEraserMenu(g);
        }

        private void DrawButton(Graphics g, ToolButton btn)
        {
            if (!_buttonRects.TryGetValue(btn, out var rect))
                return;

            bool isCurrentMode = false;
            try
            {
                switch (btn)
                {
                    case ToolButton.Arrow: isCurrentMode = !Service.IsPenMode && !Service.IsEraserMode; break;
                    case ToolButton.Pen: isCurrentMode = Service.IsPenMode; break;
                    case ToolButton.Eraser: isCurrentMode = Service.IsEraserMode; break;
                }
            }
            catch { }

            bool isHover = _hoverButton == btn;

            // ─── 背景填充（极淡白色或悬停稍亮） ───
            Color bg = isHover ? HoverBgColor : BgColor;
            using (var brush = new SolidBrush(bg))
                g.FillEllipse(brush, rect);

            // ─── 边框（白色半透明） ───
            Color borderClr;
            float borderW = BorderWidth;
            if (isCurrentMode)
            {
                borderClr = ActiveBorderColor;
                borderW = 2.5f;
            }
            else if (isHover)
            {
                borderClr = HoverBorderColor;
            }
            else
            {
                borderClr = BorderColor;
            }

            using (var pen = new Pen(borderClr, borderW))
                g.DrawEllipse(pen, rect);

            // ─── 绘制自定义图标 ───
            Color iconClr = isHover ? HoverIconColor : IconColor;
            DrawCustomIcon(g, btn, rect, iconClr);
        }

        /// <summary>绘制自定义矢量图标</summary>
        private void DrawCustomIcon(Graphics g, ToolButton btn, Rectangle rect, Color color)
        {
            int cx = rect.X + rect.Width / 2;
            int cy = rect.Y + rect.Height / 2;
            float s = rect.Width * 0.35f; // 缩放因子

            using (var pen = new Pen(color, 2.5f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
            using (var brush = new SolidBrush(color))
            {
                switch (btn)
                {
                    case ToolButton.Arrow:
                        // 箭头：一个向右的箭头
                        pen.Width = 3f;
                        g.DrawLine(pen, cx - s, cy, cx + s, cy);
                        g.DrawLine(pen, cx + s, cy, cx + s - 6, cy - 6);
                        g.DrawLine(pen, cx + s, cy, cx + s - 6, cy + 6);
                        break;

                    case ToolButton.Pen:
                        // 画笔：一支斜放的笔
                        pen.Width = 2.5f;
                        float penLen = s * 1.2f;
                        g.DrawLine(pen, cx - penLen, cy + penLen * 0.4f,
                                       cx + penLen, cy - penLen * 0.4f);
                        // 笔尖
                        float tipX = cx + penLen;
                        float tipY = cy - penLen * 0.4f;
                        g.DrawLine(pen, tipX, tipY, tipX + 4, tipY - 4);
                        // 笔尾
                        float tailX = cx - penLen;
                        float tailY = cy + penLen * 0.4f;
                        g.DrawLine(pen, tailX, tailY, tailX - 3, tailY + 3);
                        // 笔头小圆
                        g.FillEllipse(brush, tipX - 1, tipY - 1, 3, 3);
                        break;

                    case ToolButton.Eraser:
                        // 橡皮擦：一个矩形
                        float erW = s * 1.1f;
                        float erH = s * 0.7f;
                        using (var path = new GraphicsPath())
                        {
                            path.AddRectangle(new RectangleF(cx - erW / 2, cy - erH / 2, erW, erH));
                            g.DrawPath(pen, path);
                        }
                        // 擦除线条
                        using (var erasePen = new Pen(Color.FromArgb(120, color), 2f))
                        {
                            erasePen.StartCap = LineCap.Round;
                            erasePen.EndCap = LineCap.Round;
                            g.DrawLine(erasePen, cx - erW * 0.3f, cy, cx + erW * 0.3f, cy);
                        }
                        break;

                    case ToolButton.Exit:
                        // 退出：× 符号
                        pen.Width = 3f;
                        float exitS = s * 0.7f;
                        g.DrawLine(pen, cx - exitS, cy - exitS, cx + exitS, cy + exitS);
                        g.DrawLine(pen, cx + exitS, cy - exitS, cx - exitS, cy + exitS);
                        // 小方框
                        float boxS = s * 0.9f;
                        using (var path = new GraphicsPath())
                        {
                            path.AddRectangle(new RectangleF(cx - boxS, cy - boxS, boxS * 2, boxS * 2));
                            g.DrawPath(pen, path);
                        }
                        break;
                }
            }
        }

        private void DrawColorPicker(Graphics g)
        {
            if (!_buttonRects.TryGetValue(ToolButton.Pen, out var penRect))
                return;

            int cs = 24; int cols = 4, rows = 2, cp = 5;
            int panelW = cols * (cs + cp) + cp + 16;
            int panelH = rows * (cs + cp) + cp + 28;

            int px = penRect.X + BtnSize / 2 - panelW / 2;
            int py = penRect.Y - panelH - 12;
            if (py < 0) py = penRect.Bottom + 12;

            _colorPanelRect = new Rectangle(px, py, panelW, panelH);
            _colorRects.Clear();

            // 背景
            using (var bgBrush = new SolidBrush(Color.FromArgb(210, 40, 40, 40)))
            using (var path = CreateRoundRect(new Rectangle(px, py, panelW, panelH), 10))
            {
                g.FillPath(bgBrush, path);
            }

            using (var font = new Font("Segoe UI", 9, FontStyle.Bold))
            using (var brush = new SolidBrush(Color.FromArgb(200, 255, 255, 255)))
            {
                g.DrawString("画笔颜色", font, brush, px + cp + 4, py + cp + 2);
            }

            int startX = px + cp + 4;
            int startY = py + cp + 24;
            int cx = startX, cy = startY;
            for (int i = 0; i < _presetColors.Length; i++)
            {
                var cr = new Rectangle(cx, cy, cs, cs);
                _colorRects[i] = cr;

                using (var brush = new SolidBrush(_presetColors[i]))
                    g.FillEllipse(brush, cr);
                using (var borderPen = new Pen(Color.FromArgb(100, 255, 255, 255), 1))
                    g.DrawEllipse(borderPen, cr);

                if (_presetColors[i] == _penColor)
                {
                    using (var selPen = new Pen(Color.White, 2.5f))
                        g.DrawEllipse(selPen, cr);
                }

                cx += cs + cp;
                if ((i + 1) % cols == 0) { cx = startX; cy += cs + cp; }
            }

            AddHitArea(_colorPanelRect);
        }

        private void DrawEraserMenu(Graphics g)
        {
            if (!_buttonRects.TryGetValue(ToolButton.Eraser, out var eraserRect))
                return;

            int mw = 150, mh = 42;
            int mx = eraserRect.X + BtnSize / 2 - mw / 2;
            int my = eraserRect.Y - mh - 12;
            if (my < 0) my = eraserRect.Bottom + 12;

            _eraserMenuRect = new Rectangle(mx, my, mw, mh);

            using (var bgBrush = new SolidBrush(Color.FromArgb(210, 40, 40, 40)))
            using (var path = CreateRoundRect(_eraserMenuRect, 10))
            using (var borderPen = new Pen(Color.FromArgb(100, 255, 255, 255), 1))
            {
                g.FillPath(bgBrush, path);
                g.DrawPath(borderPen, path);
            }

            using (var font = new Font("Segoe UI", 11, FontStyle.Regular))
            using (var brush = new SolidBrush(Color.FromArgb(230, 255, 255, 255)))
            {
                g.DrawString("🗑  除该页笔迹", font, brush, mx + 12, my + 10);
            }

            AddHitArea(_eraserMenuRect);
        }

        private void ToolbarOverlay_MouseMove(object sender, MouseEventArgs e)
        {
            ToolButton? newHover = null;
            foreach (var kvp in _buttonRects)
            {
                if (kvp.Value.Contains(e.Location))
                { newHover = kvp.Key; break; }
            }
            if (newHover != _hoverButton)
            { _hoverButton = newHover; Invalidate(); }
        }

        private void ToolbarOverlay_MouseClick(object sender, MouseEventArgs e)
        {
            if (_showColorPicker)
            {
                bool clickedColor = false;
                foreach (var kvp in _colorRects)
                {
                    if (kvp.Value.Contains(e.Location))
                    {
                        _penColor = _presetColors[kvp.Key];
                        Service.SetPenColor(_penColor);
                        clickedColor = true;
                        break;
                    }
                }
                _showColorPicker = false;
                Invalidate();
                OnButtonClicked();
                if (clickedColor) return;
            }

            if (_showEraserMenu)
            {
                if (_eraserMenuRect.Contains(e.Location))
                    Service.EraseDrawing();
                _showEraserMenu = false;
                Invalidate();
                OnButtonClicked();
                return;
            }

            foreach (var kvp in _buttonRects)
            {
                if (!kvp.Value.Contains(e.Location)) continue;

                switch (kvp.Key)
                {
                    case ToolButton.Arrow: Service.SetPointerArrow(); break;
                    case ToolButton.Pen:
                        if (Service.IsPenMode)
                        { _showColorPicker = !_showColorPicker; _showEraserMenu = false; }
                        else
                        { Service.SetPointerPen(); _showColorPicker = false; _showEraserMenu = false; }
                        break;
                    case ToolButton.Eraser:
                        if (Service.IsEraserMode)
                        { _showEraserMenu = !_showEraserMenu; _showColorPicker = false; }
                        else
                        { Service.SetPointerEraser(); _showEraserMenu = false; _showColorPicker = false; }
                        break;
                    case ToolButton.Exit: Service.ExitShow(); break;
                }
                Invalidate();
                OnButtonClicked();
                return;
            }

            if (_showColorPicker || _showEraserMenu)
            { _showColorPicker = false; _showEraserMenu = false; Invalidate(); }
        }

        private static GraphicsPath CreateRoundRect(Rectangle rect, int r)
        {
            var path = new GraphicsPath();
            int r2 = r * 2;
            path.AddArc(rect.X, rect.Y, r2, r2, 180, 90);
            path.AddArc(rect.Right - r2, rect.Y, r2, r2, 270, 90);
            path.AddArc(rect.Right - r2, rect.Bottom - r2, r2, r2, 0, 90);
            path.AddArc(rect.X, rect.Bottom - r2, r2, r2, 90, 90);
            path.CloseFigure();
            return path;
        }
    }
}
