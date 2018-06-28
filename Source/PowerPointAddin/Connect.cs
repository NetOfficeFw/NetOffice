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
    [ProgId("PPAddin.Connect"), Guid("B6A2376C-1C4A-4917-B5DA-01442CF2C71F"), Codebase, Timestamp]
    [CustomUI("RibbonUI.xml", true)]
    [CustomPane(typeof(PowerPointAddin.Pane), "Source", false, PaneDockPosition.msoCTPDockPositionTop, PaneDockPositionRestrict.msoCTPDockPositionRestrictNoVertical, 60, 60)]
    public class Connect : COMAddin
    {
        public Connect()
        {

        }

        protected override Core CreateFactory()
        {
            var factory =  base.CreateFactory();
            factory.ObjectActivator.RegisterType(typeof(Office.ICTPFactory), typeof(MyICTPFactory));
            factory.ObjectActivator.RegisterType(typeof(Office.CustomTaskPane), typeof(MyCustomTaskPane));
            return factory;
        }

        public override void CTPFactoryAvailable(object CTPFactoryInst)
        {
            try
            {
               base.CTPFactoryAvailable(CTPFactoryInst);
            }
            catch (Exception exception)
            {
                throw;
            }           
        }

        protected override void RaiseOnStartupComplete(ref Array custom)
        {
            base.RaiseOnStartupComplete(ref custom);
        }

        protected override void RaiseOnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            base.RaiseOnConnection(Application, ConnectMode, AddInInst, ref custom);
        }

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
