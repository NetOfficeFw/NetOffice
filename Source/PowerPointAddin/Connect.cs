using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Tools;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using PowerPoint = NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Enums;
using NetOffice.PowerPointApi.Tools;

namespace PowerPointAddin
{
    [COMAddin("Addin Source Sample Addin CS4", "Addin Source Sample", LoadBehavior.LoadAtStartup)]
    [ProgId("PowerPointAddin.Connect"), Guid("A922E452-D598-4CF4-9239-B23093F17783"), Codebase, Timestamp]
    [RegistryLocation(RegistrySaveLocation.InstallScopeCurrentUser)]
    [CustomUI("RibbonUI.xml", true)]
    [CustomPane(typeof(Pane), "Source", false, PaneDockPosition.msoCTPDockPositionTop, PaneDockPositionRestrict.msoCTPDockPositionRestrictNoVertical, 60, 60)]
    public class Connect : COMAddin
    {
        protected override void TaskPaneVisibleStateChanged(Office._CustomTaskPane customTaskPaneInst)
        {
            if (null != RibbonUI)
                RibbonUI.InvalidateControl("PaneVisibleToogleButton");
        }

        public bool OnGetPressedPanelToggle(Office.IRibbonControl control)
        {
            if (TaskPanes.Count > 0)
                return TaskPanes[0].Visible;
            else
                return false;
        }

        public void OnCheckPanelToggle(Office.IRibbonControl control, bool pressed)
        {
            if (TaskPanes.Count > 0)
                TaskPanes[0].Visible = pressed;
        }

        public void OnClickAboutButton(Office.IRibbonControl control)
        {
            Utils.Dialog.ShowDiagnostics();
        }
    }
}
