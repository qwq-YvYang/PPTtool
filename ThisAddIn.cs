using System;
using System.Windows.Forms;
using Microsoft.Office.Tools;

namespace PPTTool
{
    public partial class ThisAddIn
    {
        private SlideShowTracker _tracker;
        private static ThisAddIn _instance;

        /// <summary>全局实例，供其他类访问 Application 对象</summary>
        internal static ThisAddIn Instance => _instance;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _instance = this;

            // 创建放映跟踪器，挂钩事件
            _tracker = new SlideShowTracker(Application);

            System.Diagnostics.Debug.WriteLine("[PPTTool] 外接程序已启动");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // 清理资源
            if (_tracker != null)
            {
                _tracker.Dispose();
                _tracker = null;
            }

            _instance = null;

            System.Diagnostics.Debug.WriteLine("[PPTTool] 外接程序已关闭");
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
