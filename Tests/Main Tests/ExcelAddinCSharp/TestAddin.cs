using System;
using System.Runtime.InteropServices;

using NetOffice.Tools;
using Excel = NetOffice.ExcelApi;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using NetOffice.ExcelApi.Tools;

namespace ExcelAddinCSharp
{
    [COMAddin("NOTestsMain.ExcelTestAddinCSharp", "This is a test addin from NOTests.Main", 3), Tweak(true)]
    [Guid("D48A7B31-8C03-43A8-8504-3883843799A8"), ProgId("NOTestsMain.ExcelTestAddinCSharp"), CustomUI("ExcelAddinCSharp.RibbonUI.xml")]
    public class TestAddin : COMAddin
    {
        #region Ctor
        
        public TestAddin()
        {
            Factory.Console.Name = "ExcelTestAddinCSharp";
            TaskPanes.Add(typeof(SampleControl), "NOTestsMain - C# Excel Pane");
            TaskPanes[0].DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            TaskPanes[0].DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal;
            TaskPanes[0].Width = 150;
            TaskPanes[0].Visible = true;
            TaskPanes[0].Arguments = new object[] { this };
            TaskPanes[0].Arguments = new object[] { this };
            this.OnConnection += new OnConnectionEventHandler(TestAddin_OnConnection);
        }

        #endregion

        #region Properties
        
        public bool StatusOkay
        {
            get
            {
                if (RibbonUIOkay && TaskPaneOkay  && TweakOkay && null == GeneralError)
                    return true;
                else
                    return false;
            }
        }

        public string StatusDescription
        {
            get
            {
                string result = "";
                if (!TaskPaneOkay)
                    result += "Taskpane is not loaded";
                if (!RibbonUIOkay)
                    result += "RibbonUI is not loaded";
                if(!TweakOkay)
                    result += "Tweak is not set " + Factory.Settings.ExceptionMessage;
                if (null != GeneralError)
                    result += "General Error:" + GeneralError;

                return result;
            }
        }

        private string GeneralError { get; set; }

        internal bool RibbonUIOkay 
        {
            get
            {
                return null != RibbonUI;
            }
        }
        
        internal bool TweakOkay
        {
            get
            {
                if (Factory.Settings.ExceptionMessage.StartsWith ("Test09TweakCS"))
                    return true;
                else
                    return false;
            }
        }

        internal bool TaskPaneOkay { get; set; }

        internal Office.IRibbonUI RibbonUI { get; private set; }

        #endregion

        #region Trigger
        
        private void TestAddin_OnConnection(object Application, NetOffice.Tools.ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            Factory.Initialize();
            Office.COMAddIn addin = new Office.COMAddIn(null, AddInInst);
            addin.Object = this;
            addin.Dispose();
        }

        public void OnLoadRibbonUI(Office.IRibbonUI ribbonUI)
        {
            RibbonUI = ribbonUI;
        }

        public string GetLabel(Office.IRibbonControl control)
        {
            return Factory.Settings.ExceptionMessage;
        }

        #endregion

        #region Overrides

        protected override bool AllowApplyTweak(string name, string value)
        {
            Factory.Console.SendPipeConsoleMessage("ExcelTestAddinCSharp", String.Format("AllowApplyTweak {0}:{1}", name, value));
            return true;
        }

        protected override void OnError(ErrorMethodKind methodKind, Exception exception)
        {
            if (null == GeneralError)
                GeneralError = "";
            GeneralError += methodKind.ToString() + Environment.NewLine + exception.GetType().Name + Environment.NewLine + exception.Message;

        }

        [RegisterFunction(RegisterMode.CallAfter)]
        public static void Register(Type type, RegisterCall registerCall)
        {
            SetTweakPersistenceEntry(type, "NOExceptionMessage", "Test09TweakCS", false);
        }

        #endregion
    }
}
