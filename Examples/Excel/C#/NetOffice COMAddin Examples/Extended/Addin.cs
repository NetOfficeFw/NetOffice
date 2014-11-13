using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using NetOffice.Tools;
using NetOffice.OfficeApi.Enums;
using NetOffice.ExcelApi.Tools;
using Excel = NetOffice.ExcelApi;
using Office = NetOffice.OfficeApi;

/*
  * This project shows you in depth the COMAddin base class from the NetOffice tools.
  * The COMAddin base class is designed to reduce infrastructure code.
  * This addin looks a bit strange of course because the explanation.
  * Check the NetOffice download section for more NetOffice Tools based Addins
  * Wikipedia Addin  - Word
  * Twitter Addin    - Outlook
  * Google Addin     - Excel
*/

namespace NetOfficeTools.ExtendedExcelCS4
{
    /*
     * As you can see, the necessary registry informations was given as annotation, no need for Register/Unregister methods
     * The RegistryLocation attribute is not always necessary. CurrentUser is default, no need for this attribute if you dont want HKEY_CURRENTUSER (just for example here)
     * You see also the CustomUI attribute. You can specify a path to an embedded xml ressource file with your ribbon schema. If you dont want this then you can override the GetCustomUI method from the base class.
     * The Tweak attribute allows to set various NetOffice options at runtime with custom values entries in the current office addin key(helpful for troubleshooting). Learn more about in the Tweaks sample addin project.
     * The CustomPane attribute allows you to set a task pane very easy
     */
    [COMAddin("NetOfficeCS4 Extended Sample Addin", "This Addin shows you the COMAddin class from the NetOffice Tools", 3)]
    [CustomUI("NetOfficeTools.ExtendedExcelCS4.RibbonUI.xml")]
    [RegistryLocation(RegistrySaveLocation.CurrentUser)]
    [CustomPane(typeof(SamplePane), "NetOffice Tools - Sample Pane(CS4)", true, PaneDockPosition.msoCTPDockPositionTop, PaneDockPositionRestrict.msoCTPDockPositionRestrictNoChange, 50, 50)]
    [Guid("BA38FD48-47BD-43de-8177-0D067A01B566"), ProgId("ExtendedExcelCS4.Addin"), Tweak(true)]
    public class Addin : COMAddin
    {
        public Addin()
        {
            // trigger the well known IExtensibility2 methods, this is very similar to VSTO
            this.OnStartupComplete += new OnStartupCompleteEventHandler(Addin_OnStartupComplete);

            /*
            // We add a second taskpane here at hand, you can also overwrite the CTPFactoryAvailable method and create your panes in this method.
            // Taskpanes in Netoffice can implement the ITaskPane interface with the OnConnection/OnDisconnection to avoid the singleton pattern.
            // Take a look into the SamplePane.cs to see how you can use the NetOffice ITaskPane interface to get more control for Load/Unload and connect the host application.
            
            Office.Tools.TaskPaneInfo pane = TaskPanes.Add(typeof(SamplePane), "NetOffice Tools - 2. Sample Pane(CS4)");
            pane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionTop;
            pane.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
            pane.Height = 50;
            pane.Visible = true;
            pane.Arguments = new object[] { this };
            pane.VisibleStateChange += new Office.CustomTaskPane_VisibleStateChangeEventHandler(TaskPane_VisibleStateChange);
            */
        }

        // ouer ribbon instance to manipulate ui at runtime 
        internal Office.IRibbonUI RibbonUI { get; private set; }

        // attached in ctor to say hello in console
        private void Addin_OnStartupComplete(ref Array custom)
        {
            // Tou see the host application is accessible as property from the class instance.
            // The application property was disposed automaticly while shutdown.
            Factory.Console.WriteLine("Host Application Version is:{0}", this.Application.Version);
        }

        //// attached in ctor to trigger taskpane visibility has been changed and update the checkbutton in the ribbon ui for show/hide taskpane
        //private void TaskPane_VisibleStateChange(NetOffice.OfficeApi._CustomTaskPane CustomTaskPaneInst)
        //{           
        //    // ouer taskpane visibility has been changed. we send a message to the host application
        //    // and say please refresh the checkbutton state. now the host application want call ouer OnGetPressedPanelToggle method to update the checkstate.
        //    RibbonUI.InvalidateControl("paneVisibleToogleButton");
        //}

        // taskpane visibility has been changed. we upate the checkbutton in the ribbon ui for show/hide taskpane
        protected override void TaskPaneVisibleStateChanged(Office._CustomTaskPane customTaskPaneInst)
        {
            if (null != RibbonUI)
                RibbonUI.InvalidateControl("paneVisibleToogleButton");
        }

        // defined in RibbonUI.xml to get a instance for ouer ribbon ui.
        public void OnLoadRibonUI(Office.IRibbonUI ribbonUI)
        {
            RibbonUI = ribbonUI;
        }

        // defined in RibbonUI.xml to make sure the checkbutton state is up-to-date and synchronized with taskpane visibility.
        public bool OnGetPressedPanelToggle(Office.IRibbonControl control)
        {
            return TaskPanes[0].Visible;
        }

        // defined in RibbonUI.xml to track the user clicked ouer checkbutton. then we upate the panel visibility at hand.
        public void OnCheckPanelToggle(Office.IRibbonControl control, bool pressed)
        {
            TaskPanes[0].Visible = pressed;
        }

        // defined in RibbonUI.xml to catch the user click for the about button
        public void OnClickAboutButton(Office.IRibbonControl control)
        {
            MessageBox.Show("NetOffice Tools - Extended Sample Addin.", "ExtendedExcelCS4.Addin");
        }

        /*
        * Now you see the way to exend or modify the register/unregister process if you want.
        * We define 2 static methods with the RegisterFunction attribute, we use CallBeforeAndAfter as condition.
        * This condition means the register method in the base class call our method as first (before registry modification) and as last(after registry modification).
        * The register call argument give you the info what is it currently. Replace means the method in the base class does nothing and its your task to create the registry keys.
        * Same thing with Unregister method. 
        */

        [RegisterFunction(RegisterMode.CallBeforeAndAfter)]
        public static void Register(Type type, RegisterCall registerCall)
        {
            switch (registerCall)
            {
                case RegisterCall.CallAfter:
                    break;
                case RegisterCall.CallBefore:
                    break;
                case RegisterCall.Replace:
                    break;
                default:
                    break;
            }
        }

        [UnRegisterFunction(RegisterMode.CallBeforeAndAfter)]
        public static void UnRegister(Type type, RegisterCall registerCall)
        {
            switch (registerCall)
            {
                case RegisterCall.CallAfter:
                    break;
                case RegisterCall.CallBefore:
                    break;
                case RegisterCall.Replace:
                    break;
                default:
                    break;
            }
        }


        /*
         * at last you see some options for troubleshooting. the COMAddin base class is not a blackbox.
        */

        // This error handler is used for IExtensibility2 events (your code) and the COMAddin methods GetCustomUI, CTPFactoryAvailable and CreateFactory(also overwrites).
        // the first argument shows in which method the error is occured. The second argument is the detailed exception info. 
        // Rethrow the exception otherwise the exception is marked as handled.   
        protected override void OnError(ErrorMethodKind methodKind, Exception exception)
        {
            MessageBox.Show("An error occurend in " + methodKind.ToString(), "ExtendedExcelCS4.Addin");
        }

        // This method demonstrate an error handler for the register/unregister process.
        // For example you have an security issues while register or something like that then you can implement a static errorhandler method.
        // The first argument shows you the error occurs in Register or Unregister.
        // The second argument is the thrown exception. Rethrow the exception to signalize an error to the environment otherwise the exception is handled.
        [RegisterErrorHandler]
        public static void RegisterErrorHandler(RegisterErrorMethodKind methodKind, Exception exception)
        {
            MessageBox.Show("An register error occurend in " + methodKind.ToString(), "ExtendedExcelCS4.Addin");
        }
    }
}
