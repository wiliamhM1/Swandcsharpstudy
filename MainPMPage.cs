using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using System;
using System.Runtime.InteropServices;

namespace SwCSharpAddin2
{
    [ComVisible(true)]
    public class MainPMPage : PropertyManagerPage2Handler9
    {
        private readonly SwAddin _addin;
        private IPropertyManagerPage2 _page;

        // 模块分组
        public IPropertyManagerPageGroup GroupModelSetup { get; private set; }
        public IPropertyManagerPageGroup GroupTrajectory { get; private set; }
        public IPropertyManagerPageGroup GroupCodeProcessing { get; private set; }

        public MainPMPage(SwAddin addin)
        {
            _addin = addin;
            CreatePage();
        }

        private void CreatePage()
        {
            try
            {
                const string PAGE_TITLE = "模型处理器";
                const swPropertyManagerPageOptions_e OPTIONS =
                    swPropertyManagerPageOptions_e.swPropertyManagerOptions_OkayButton |
                    swPropertyManagerPageOptions_e.swPropertyManagerOptions_CancelButton;

                int error = 0;
                _page = _addin.SwApp.CreatePropertyManagerPage(
                    PAGE_TITLE,
                    (int)OPTIONS,
                    this,
                    ref error) as IPropertyManagerPage2;

                if (_page == null || error != 0)
                {
                    MessageBox.Show($"创建属性页失败，错误代码: {error}");
                    return;
                }

                // ====== 1. 模型设置模块 ======
                GroupModelSetup = _page.AddGroupBox(
                    id: 1,
                    caption: "模型设置",
                    options: (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Expanded
                );

                GroupModelSetup.AddControl(
                    id: 10,
                    type: (int)swPropertyManagerPageControlType_e.swControlType_Label,
                    text: "配置几何参数和约束条件",
                    options: 0,
                    leftAlign: true
                );

                // ====== 2. 轨迹提取模块 ======
                GroupTrajectory = _page.AddGroupBox(
                    id: 2,
                    caption: "轨迹提取",
                    options: (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Expanded
                );

                GroupTrajectory.AddControl(
                    id: 20,
                    type: (int)swPropertyManagerPageControlType_e.swControlType_Label,
                    text: "定义运动路径和特征点",
                    options: 0,
                    leftAlign: true
                );

                // ====== 3. 代码处理模块 ======
                GroupCodeProcessing = _page.AddGroupBox(
                    id: 3,
                    caption: "代码处理",
                    options: (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Expanded
                );

                GroupCodeProcessing.AddControl(
                    id: 30,
                    type: (int)swPropertyManagerPageControlType_e.swControlType_Label,
                    text: "生成控制指令和加工代码",
                    options: 0,
                    leftAlign: true
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show($"界面创建错误: {ex.Message}");
            }
        }

        public void Show() => _page?.Show();

        public void Close()
        {
            if (_page != null)
            {
                Marshal.ReleaseComObject(_page);
                _page = null;
            }
        }

        #region 事件处理
        public override void AfterClose(int reason)
        {
            // 清理资源但不销毁实例
            Close();
        }
        #endregion 
    }
}