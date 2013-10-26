using System;
using System.Runtime.InteropServices;
using Access = NetOffice.AccessApi;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using NetOffice.AccessApi.Tools;
using NetOffice.Tools;

namespace TestAddin
{
    [COMAddin("TestAddin.Addin", "Test Addin", 3)]
    [Guid("FC790F5D-C7CD-4B0E-A273-50E7370F628D"), ProgId("TestAddin.Addin"), Tweak(true)]  
    public class Addin : COMAddin
    {
        public Addin()
        {
            // we can add our own taskpanes here, if you dont want that then overwrite the CTPFactoryAvailable method
            // show into the SamplePane.cs to see how you can use the NetOffice ITaskPane interface to get more control for Load/Unload and connect the host application
            TaskPanes.Add(typeof(TestPane), "NetOffice Tools - Sample Pane(CS4)");
            TaskPanes[0].DockPosition = MsoCTPDockPosition.msoCTPDockPositionTop;
            TaskPanes[0].DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
            TaskPanes[0].Height = 50;
            TaskPanes[0].Visible = true;
            TaskPanes[0].Arguments = new object[] { this };
            TaskPanes[0].VisibleStateChange += new Office.CustomTaskPane_VisibleStateChangeEventHandler(TaskPane_VisibleStateChange);
            this.OnStartupComplete += new OnStartupCompleteEventHandler(Addin_OnStartupComplete);
        }

        private void Addin_OnStartupComplete(ref Array custom)
        {
        }

        private void TaskPane_VisibleStateChange(Office._CustomTaskPane CustomTaskPaneInst)
        {
            
        }

        protected override bool AllowApplyTweak(string name, string value)
        {
            return true;
        }

        protected override void ApplyCustomTweak(string name, string value)
        {
            if (name == "ShowMessageBoxAtStartUp" && value == "yes")
                System.Windows.Forms.MessageBox.Show("The tweak sample addin has been loaded.", "Custom Tweak", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        }

        protected override void DisposeCustomTweak(string name, string value)
        {

        }

        [RegisterFunction(RegisterMode.CallAfter)]
        public static void Register(Type type, RegisterCall registerCall)
        {
            SetTweakPersistenceEntry(type, "NOExceptionMessage", "TestMessage", false);
            SetTweakPersistenceEntry(type, "NOConsoleMode", "trace", false);
            SetTweakPersistenceEntry(type, "ShowMessageBoxAtStartUp", "yes", false);
        }
    }
}
