using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace PPTTool
{
    /// <summary>
    /// 翻页栏通用类（白色半透明圆形边框 + 矢量图标）
    /// 每个翻页栏包含：上一页（上）+ 下一页（下）两个圆形按钮
    /// </summary>
    public class NavOverlay : OverlayFormBase
    {
        public enum NavPosition { Left, Right }

        private const int BtnSize = 48;
        private const int BtnGap = 8;
        private const int EdgePad = 14;
        private const int BorderWidth = 2;

        private readonly NavPosition _position;
        private Rectangle _prevBtnRect;
        private Rectangle _nextBtnRect;
        private bool _hoverPrev;
        private bool _hoverNext;

        // ─── UI 颜色配置（后期可自定义） ───
        public Color BorderColor { get; set; } = Color.FromArgb(160, 255, 255, 255);
        public Color BgColor { get; set; } = Color.FromArgb(30, 255, 255, 255);
        public Color HoverBorderColor { get; set; } = Color.FromArgb(220, 255, 255, 255);
        public Color HoverBgColor { get; set; } = Color.FromArgb(80, 255, 255, 255);
        public Color IconColor { get; set; } = Color.FromArgb(200, 255, 255, 255);
        public Color HoverIconColor { get; set; } = Color.FromArgb(255, 255, 255, 255);

        public NavOverlay(PowerPointService service, NavPosition position) : base(service)
        {
            _position = position;

            int contentHeight = BtnSize * 2 + BtnGap;
            int contentWidth = BtnSize;
            Width = contentWidth + EdgePad * 2;
            Height = contentHeight + EdgePad * 2;

            var screen = Screen.PrimaryScreen.WorkingArea;

            if (position == NavPosition.Left)
                Location = new Point(20, screen.Height / 2 + 40);
            else
                Location = new Point(screen.Width - Width - 20, screen.Height / 2 + 40);

            DoubleBuffered = true;
            Paint += NavOverlay_Paint;
            MouseClick += NavOverlay_MouseClick;
            MouseMove += NavOverlay_MouseMove;
            MouseLeave += (s, e) =>
            {
                if (_hoverPrev || _hoverNext)
                { _hoverPrev = false; _hoverNext = false; Invalidate(); }
            };

            LayoutButtons();
        }

        private void LayoutButtons()
        {
            ClearHitAreas();
            int cx = EdgePad;
            int cy = EdgePad;

            _prevBtnRect = new Rectangle(cx, cy, BtnSize, BtnSize);
            AddHitArea(_prevBtnRect);

            cy += BtnSize + BtnGap;
            _nextBtnRect = new Rectangle(cx, cy, BtnSize, BtnSize);
            AddHitArea(_nextBtnRect);
        }

        private void NavOverlay_Paint(object sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;

            DrawNavButton(g, _prevBtnRect, true, _hoverPrev);
            DrawNavButton(g, _nextBtnRect, false, _hoverNext);
        }

        private void DrawNavButton(Graphics g, Rectangle rect, bool isPrev, bool hover)
        {
            Color bg = hover ? HoverBgColor : BgColor;
            using (var brush = new SolidBrush(bg))
                g.FillEllipse(brush, rect);

            Color borderClr = hover ? HoverBorderColor : BorderColor;
            float bw = hover ? 2.5f : BorderWidth;
            using (var pen = new Pen(borderClr, bw))
                g.DrawEllipse(pen, rect);

            // 绘制矢量箭头
            Color iconClr = hover ? HoverIconColor : IconColor;
            DrawArrowIcon(g, rect, isPrev, iconClr);
        }

        /// <summary>绘制三角箭头矢量图标</summary>
        private void DrawArrowIcon(Graphics g, Rectangle rect, bool isPrev, Color color)
        {
            int cx = rect.X + rect.Width / 2;
            int cy = rect.Y + rect.Height / 2;
            float size = rect.Width * 0.3f;

            // 箭头方向：左（上一页）/ 右（下一页）
            float dir = isPrev ? -1 : 1;

            using (var pen = new Pen(color, 3f)
            {
                StartCap = LineCap.Round,
                EndCap = LineCap.Round
            })
            {
                // 水平线（箭头杆）
                g.DrawLine(pen, cx - dir * size, cy, cx + dir * size, cy);

                // 箭头尖（上半部 + 下半部）
                float tipX = cx + dir * size;
                g.DrawLine(pen, tipX, cy, tipX - dir * 7, cy - 7);
                g.DrawLine(pen, tipX, cy, tipX - dir * 7, cy + 7);
            }
        }

        private void NavOverlay_MouseClick(object sender, MouseEventArgs e)
        {
            if (_prevBtnRect.Contains(e.Location))
            {
                Service.PreviousSlide();
                OnButtonClicked();
            }
            else if (_nextBtnRect.Contains(e.Location))
            {
                Service.NextSlide();
                OnButtonClicked();
            }
        }

        private void NavOverlay_MouseMove(object sender, MouseEventArgs e)
        {
            bool newHoverPrev = _prevBtnRect.Contains(e.Location);
            bool newHoverNext = _nextBtnRect.Contains(e.Location);

            if (newHoverPrev != _hoverPrev || newHoverNext != _hoverNext)
            {
                _hoverPrev = newHoverPrev;
                _hoverNext = newHoverNext;
                Invalidate();
            }
        }
    }
}
