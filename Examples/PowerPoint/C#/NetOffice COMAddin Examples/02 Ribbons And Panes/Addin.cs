using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Tools;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using Excel = NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Enums;
using NetOffice.PowerPointApi.Tools;
/*
    Ribbons & Panes Addin Example
*/
namespace PowerPoint02AddinCS4
{
    [COMAddin("PowerPoint02 Sample Addin CS4", "Ribbons & Panes Addin Example", LoadBehavior.LoadAtStartup)]
    [ProgId("PowerPoint02AddinCS4.Connect"), Guid("A8589603-587D-4887-9DE6-19F03C53304A"), Codebase, Timestamp]
    [CustomUI("RibbonUI.xml", true)]
    [CustomPane(typeof(SamplePane), "PowerPoint CPU Usage", false, PaneDockPosition.msoCTPDockPositionTop, PaneDockPositionRestrict.msoCTPDockPositionRestrictNoVertical, 60, 60)]   
    public class Addin : COMAddin
    { 
        // Taskpane visibility has been changed. We upate the checkbutton in the ribbon ui for show/hide taskpane
        protected override void TaskPaneVisibleStateChanged(Office._CustomTaskPane customTaskPaneInst)
        {
            if (null != RibbonUI)
                RibbonUI.InvalidateControl("PaneVisibleToogleButton");
        }

        // Defined in RibbonUI.xml to make sure the checkbutton state is up-to-date and synchronized with taskpane visibility.
        public bool OnGetPressedPanelToggle(Office.IRibbonControl control)
        {
            if (TaskPanes.Count > 0)
                return TaskPanes[0].Visible;
            else
                return false;
        }

        // Defined in RibbonUI.xml to track the user clicked ouer checkbutton. Then we upate the panel visibility at hand.
        public void OnCheckPanelToggle(Office.IRibbonControl control, bool pressed)
        {
            if (TaskPanes.Count > 0)
                TaskPanes[0].Visible = pressed;
        }

        // Defined in RibbonUI.xml to catch the user click for the about button
        public void OnClickAboutButton(Office.IRibbonControl control)
        {
            Utils.Dialog.ShowDiagnostics();
        }
    }
}