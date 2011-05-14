using System;
using System.Reflection; 
using System.Windows.Forms;
using System.Collections.Generic;
using Microsoft.Win32;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;

using Extensibility;

namespace SuperAddin
{ 
    /// <summary>
    /// the addin class
    /// </summary>
    [ComVisible(true)]
    [GuidAttribute("90AF5E73-9671-4c96-8E96-CC01B3DB3886"), ProgId("SuperAddin.Connect")]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private static readonly string _prodId = "SuperAddin.Connect";

        #region Fields 

        private HostApplication _application;
        private AddinUI         _userInterface;
        private TrayIcon        _trayIcon;
         
        #endregion

        #region COM Register Functions

        /// <summary>
        /// This function was called while register addin for example with regsvr32 or while compiling
        /// </summary>
        /// <param name="type"></param>
        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type type)
        {
            try
            {
                // add codebase entry
                Assembly thisAssembly = Assembly.GetAssembly(typeof(Connect));
                RegistryKey key = Registry.ClassesRoot.CreateSubKey("CLSID\\{90AF5E73-9671-4c96-8E96-CC01B3DB3886}\\InprocServer32\\1.0.0.0");
                key.SetValue("CodeBase", thisAssembly.CodeBase);
                key.Close();

                key = Registry.ClassesRoot.CreateSubKey("CLSID\\{90AF5E73-9671-4c96-8E96-CC01B3DB3886}\\InprocServer32");
                key.SetValue("CodeBase", thisAssembly.CodeBase);
                key.Close();

                // add bypass key
                // http://support.microsoft.com/kb/948461
                key = Registry.ClassesRoot.CreateSubKey("Interface\\{000C0601-0000-0000-C000-000000000046}");
                string defaultValue = key.GetValue("") as string;
                if (null == defaultValue)
                    key.SetValue("", "Office .NET Framework Lockback Bypass Key");
                key.Close();

                Registry.ClassesRoot.CreateSubKey(@"CLSID\{" + type.GUID.ToString().ToUpper() + @"}\Programmable");

                OfficeRegistry.CreateAddinKey(OfficeRegistry.Excel      + _prodId);
                OfficeRegistry.CreateAddinKey(OfficeRegistry.Word       + _prodId);
                OfficeRegistry.CreateAddinKey(OfficeRegistry.Outlook    + _prodId);
                OfficeRegistry.CreateAddinKey(OfficeRegistry.PowerPoint + _prodId);
                OfficeRegistry.CreateAddinKey(OfficeRegistry.Access     + _prodId);
            }
            catch (Exception throwedException)
            {
                FormShowError errorForm = new FormShowError("An error ocurred while register COM Addin", 
                                                            throwedException.Message, throwedException);
                errorForm.ShowDialog();
            }
        }

        /// <summary>
        /// This function called while unregister addin
        /// </summary>
        /// <param name="type"></param>
        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type type)
        {
            try
            {
                Registry.ClassesRoot.DeleteSubKey(@"CLSID\{" + type.GUID.ToString().ToUpper() + @"}\Programmable", false);
                
                OfficeRegistry.DeleteAddinKey(OfficeRegistry.Excel      + _prodId);
                OfficeRegistry.DeleteAddinKey(OfficeRegistry.Word       + _prodId);
                OfficeRegistry.DeleteAddinKey(OfficeRegistry.Outlook    + _prodId);
                OfficeRegistry.DeleteAddinKey(OfficeRegistry.PowerPoint + _prodId);
                OfficeRegistry.DeleteAddinKey(OfficeRegistry.Access     + _prodId);
            }
            catch (ArgumentException)
            {
                // key is already deleted
                ;
            }
            catch (Exception throwedException)
            {
                FormShowError errorForm = new FormShowError("An error ocurred while unregister COM Addin", 
                                                             throwedException.Message, throwedException);
                errorForm.ShowDialog();
            }
        }
        
        #endregion

        #region IDTExtensibility2 Members/Trigger

        /// <summary>
        /// The OnAddInsUpdate event occurs when the set of loaded COM add-ins changes. 
        /// When an add-in is loaded or unloaded, the OnAddInsUpdate event occurs in any other loaded add-ins. 
        /// For example, if add-ins A and B both are loaded currently, and then add-in C is loaded, 
        /// the OnAddInsUpdate event occurs in add-ins A and B. If C is unloaded, 
        /// the OnAddInsUpdate event occurs again in add-ins A and B.
        /// </summary>
        /// <param name="custom"></param>
        public void OnAddInsUpdate(ref Array custom)
        {
            try
            {
            
            }
            catch (Exception throwedException)
            {
                FormShowError.LogError("An error ocurred while perform OnAddInsUpdate.", throwedException);                    
            }
        }
        
        /// <summary>
        /// The OnBeginShutdown event occurs when the host application begins its shutdown routines, 
        /// in the case where the application closes while the COM add-in is still loaded. 
        /// If the add-in is not loaded when the application closes, the OnBeginShutdown event does not occur.
        /// When this event does occur, it occurs before the OnDisconnection event.
        /// </summary>
        /// <param name="custom"></param>
        public void OnBeginShutdown(ref Array custom)
        {
            try
            {

            }
            catch (Exception throwedException)
            {
                FormShowError.LogError("An error ocurred while perform OnBeginShutdown.", throwedException);                    
            }
        }

        /// <summary>
        /// The OnConnection event occurs when the COM add-in is loaded (connected). 
        /// </summary>
        /// <param name="Application">Provides a reference to the application in which the COM add-in is currently running. </param>
        /// <param name="ConnectMode">A constant that specifies how the add-in was loaded.</param>
        /// <param name="AddInInst">A COMAddIn object that refers to the instance of the class module in which code is currently running. You can use this argument to return the programmatic identifier for the add-in.</param>
        /// <param name="custom">An array of Variant type values that provides additional data. The numeric value of the first element in this array indicates how the host application was started: from the user interface (1), by embedding a document created in the host application in another application (2), or through Automation (3).</param>
        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                _application = new HostApplication(Application, ConnectMode, AddInInst, ref custom);
                
                _trayIcon = new TrayIcon(true);
            }
            catch (Exception throwedException)
            {
                FormShowError.LogError("An error ocurred while perform OnConnection.", throwedException);
            }
        }

         
        /// <summary>
        /// The OnDisconnection event occurs when the COM add-in is unloaded.
        /// </summary>
        /// <param name="RemoveMode">A constant that specifies how the add-in was unloaded.</param>
        /// <param name="custom">An array of Variant type values that provides additional data. The numeric value of the first element in this array indicates how the host application was started: from the user interface (1); by embedding a document created in the host application in another application (2); or through Automation (3).</param>
        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                if(null != _application)
                    _application.Dispose();

                if (null != _trayIcon)
                    _trayIcon.Dispose();
            }
            catch (Exception throwedException)
            {
                FormShowError.LogError("An error ocurred while perform OnDisconnection.", throwedException);
            }
        }

        /// <summary>
        /// The OnStartupComplete event occurs when the host application completes its startup routines, in the case where the COM add-in loads at startup.
        /// If the add-in is not loaded when the application loads, the OnStartupComplete event does not occur — even when the user loads the add-in in the COM Add-ins dialog box. When this event does occur, it occurs after the OnConnection event.
        /// </summary>
        /// <param name="custom"></param>
        public void OnStartupComplete(ref Array custom)
        {
            try
            {
                // addin UI
                _userInterface = new AddinUI(_application);
                _userInterface.ButtonClick += new ButtonClickEventHandler(_userInterface_ButtonClick);

                //excel and word document events
                _application.BeforeOpen += new OpenHandler(_application_BeforeOpen);
                _application.BeforeClose += new BeforeCloseHandler(_application_BeforeClose);
                _application.BeforeSave += new BeforeSaveHandler(_application_BeforeSave);
                _application.BeforePrint += new BeforePrintHandler(_application_BeforePrint); 

                if((null != _userInterface) && (false == _userInterface.RibbonIsActive))
                    _userInterface.ClassicUI.CreateUI(); 
            }
            catch (Exception throwedException)
            {
                FormShowError.LogError("An error ocurred while perform OnStartupComplete.", throwedException);
            }
        }

        #endregion
        
        #region IRibbonExtensibility Members

        public string GetCustomUI(string RibbonID)
        {
            try
            {
                return _userInterface.RibbonUI.GetCustomUI(RibbonID);
            }
            catch (Exception throwedException)
            {
                FormShowError.LogError("An error ocurred while perform GetCustomUI.", throwedException);
                return "";
            }
        }

        public void OnAction(IRibbonControl control)
        {
            try
            {
                _userInterface.RibbonUI.OnAction(control);
            }
            catch (Exception throwedException)
            {
                FormShowError.LogError("An error ocurred while perform OnAction.", throwedException);
            }
        }

        #endregion

        void _userInterface_ButtonClick(ButtonClickArgs args)
        {
            // from ribbon
            if (null != args.RibbonControl)
            {
                string message = string.Format("Thanks for click on a Ribbon.\r\nHostApp is {0}.{1} Version:{2}",
                                                _application.Component,  _application.Name, _application.Version);
                MessageBox.Show(message, "SuperAddin", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            // from commandbarbutton
            if (null != args.ButtonControl)
            {
                string message = string.Format("Thanks for click on a button.\r\nHostApp is {0}.{1} Version:{2}",
                                                _application.Component, _application.Name, _application.Version);
                MessageBox.Show(message, "SuperAddin" ,MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            args.Dispose();
        }

        void _application_BeforePrint(BeforePrintArgs args, ref bool Cancel)
        {            
            args.Dispose();
        }

        void _application_BeforeSave(BeforeSaveArgs args, ref bool SaveAsUI, ref bool Cancel)
        {
            args.Dispose();
        }

        void _application_BeforeClose(BeforeCloseArgs args, ref bool Cancel)
        {
            args.Dispose();
        }

        void _application_BeforeOpen(OpenArgs args)
        {
            args.Dispose();
        }
    }
}
