using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using NetOffice.Tools;
using NetOffice.OfficeApi.Enums;
using NetOffice.WordApi.Tools;
using NetOffice.WordApi;

/*
  * This project shows you in depth the COMAddin base class from the NetOffice tools.
  * The COMAddin base class is designed to reduce infrastructure code from your own.
  * this addin looks a bit strange of course because the explanation
  * check the NetOffice download section for NetOffice Tools based Addins
  * Wikipedia Addin  - Word
  * Twitter Addin    - Outlook
  * Google Addin     - Excel
*/

namespace NetOfficeTools.ExtendedWordCS4
{
    /*
     * as you can see, the needed registry informations was given as annotation, no need for Register/Unregister methods
    */
    [COMAddin("NetOfficeCS4 Extended Sample Addin", "This Addin shows you the COMAddin class from the NetOffice Tools", 3)]
    [RegistryLocation(RegistrySaveLocation.CurrentUser)]          // CurrentUser is default, no need for this attribute if you want HKEY_CURRENTUSER (just for example)
    [CustomUI("NetOfficeTools.ExtendedWordCS4.RibbonUI.xml")]     // you can specify a path to an embedded xml ressource file with your ribbon schema, otherwise you can override the GetCustomUI method from COMAddin base class
    [Guid("A37BAA2D-21CC-42DC-9AC8-F6101D6DF1AE"), ProgId("ExtendedWordCS4.Addin"), Tweak(true)]
    public class Addin : COMAddin
    {
        public Addin()
        {
            // trigger the well known IExtensibility2 methods, this is very similar to VSTO
            this.OnStartupComplete += new OnStartupCompleteEventHandler(Addin_OnStartupComplete);

            // wen can add our own taskpanes here, if you dont want that then overwrite the CTPFactoryAvailable method
            // show into the SamplePane.cs to see how you can use the NetOffice ITaskPane interface to get more control for Load/Unload and connect the host application
            TaskPanes.Add(typeof(SamplePane), "NetOffice Tools - Sample Pane");
            TaskPanes[0].DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            TaskPanes[0].DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal;
            TaskPanes[0].Width = 150;
            TaskPanes[0].Visible = true;
            TaskPanes[0].Arguments = new object[] { this };
        }

        private void Addin_OnStartupComplete(ref Array custom)
        {
            // you see the host application is accessible as property from the class instance
            // the application property was disposed automaticly while shutdown
            Console.WriteLine("Host Application Version is:{0}", this.Application.Version);
        }

        public void OnAction(NetOffice.OfficeApi.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "customButton1":
                        MessageBox.Show("This is the first sample button.", "ExtendedWordCS4.Addin");
                        break;
                    case "customButton2":
                        MessageBox.Show("This is the second sample button.", "ExtendedWordCS4.Addin");
                        break;
                    default:
                        MessageBox.Show("Unkown Control Id: " + control.Id, "ExtendedWordCS4.Addin");
                        break;
                }
            }
            catch (Exception throwedException)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine);
                MessageBox.Show("An error occured in OnAction." + details, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        /*
        * now you see the way to exend or modify the register/unregister process if you want
        * we define 2 static methods with the RegisterFunction attribute, we use CallBeforeAndAfter as parameter
        * this means the register method in the base class call our method as first (before registry modification) and as last(after registry modification) 
        * the register call parameter give you the info what is is. Replace means the method in the base class does nothing and its your task to create the registry keys
        * same thing with Unregister method. 
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

        // this error handler is used for IExtensibility2 methods (your code) and the COMAddin methods GetCustomUI and CTPFactoryAvailable        
        protected override void OnError(ErrorMethodKind methodKind, Exception exception)
        {
            MessageBox.Show("An error occurend in " + methodKind.ToString(), "ExtendedWordCS4.Addin");
        }

        // for example you have an security issues while register or something like that
        // then you can implement a static errorhandler method.
        // the first parameter shows you the error occurs in Register or Unregister
        // the second parameter is the thrown exception. rethrow the exception in this method to signalize an error to the environment
        [RegisterErrorHandler]
        public static void RegisterErrorHandler(RegisterErrorMethodKind methodKind, Exception exception)
        {
            MessageBox.Show("An error occurend in " + methodKind.ToString(), "ExtendedWordCS4.Addin");
        }
    }
}
