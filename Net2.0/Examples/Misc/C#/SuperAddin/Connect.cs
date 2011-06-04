using System;
using System.Reflection; 
using System.Windows.Forms;
using System.Collections.Generic;
using Microsoft.Win32;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using Extensibility;

using Office = NetOffice.OfficeApi;
using Excel = NetOffice.ExcelApi;
using Word = NetOffice.WordApi;
using Outlook = NetOffice.OutlookApi;
using PowerPoint = NetOffice.PowerPointApi;
using Access = NetOffice.AccessApi;

namespace SuperAddinCSharp
{ 
    /// <summary>
    /// the addin class
    /// </summary>
    [ComVisible(true)]
    [GuidAttribute("8ED7D7E2-D084-4ba7-999E-5657147460DD"), ProgId("SuperAddinCSharp.Connect")]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private static readonly string _addinName = "SuperAddinCSharp";
        private static readonly string _prodId = "SuperAddinCSharp.Connect";

        #region Fields 

        private HostApplication _application;
        private TrayIcon        _trayIcon;
        
        private bool            _isRibbonSupported;
 
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
                RegistryKey key = Registry.ClassesRoot.CreateSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\InprocServer32\\1.0.0.0");
                key.SetValue("CodeBase", thisAssembly.CodeBase);
                key.Close();

                key = Registry.ClassesRoot.CreateSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\InprocServer32");
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

                OfficeRegistry.CreateAddinKey(_addinName, OfficeRegistry.Excel + _prodId);
                OfficeRegistry.CreateAddinKey(_addinName, OfficeRegistry.Word + _prodId);
                OfficeRegistry.CreateAddinKey(_addinName, OfficeRegistry.Outlook + _prodId);
                OfficeRegistry.CreateAddinKey(_addinName, OfficeRegistry.PowerPoint + _prodId);
                OfficeRegistry.CreateAddinKey(_addinName, OfficeRegistry.Access + _prodId);
            }
            catch (Exception throwedException)
            {
                FormShowError errorForm = new FormShowError("An error ocurred while register " + _addinName, 
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
                FormShowError errorForm = new FormShowError("An error ocurred while unregister " + _addinName, 
                                                             throwedException.Message, throwedException);
                errorForm.ShowDialog();
            }
        }
        
        #endregion

        #region IDTExtensibility2 Members

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
                LateBindingApi.Core.Factory.Initialize();

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
                if(! _isRibbonSupported)
                    CreateClassicUI();
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
                _isRibbonSupported = true;
                return ReadTextFileFromRessource("RibbonUI.xml");
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
                string message = string.Format("Thanks for click on a Ribbon.\r\nHostApp is {0}.{1} Version:{2}",
                                             _application.ComponentName, _application.Name, _application.Version);
                MessageBox.Show(message, "SuperAddin", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception throwedException)
            {
                FormShowError.LogError("An error ocurred while perform OnAction.", throwedException);
            }
        }

        #endregion

        #region Classic UI

        void commandBarBtn_ClickEvent(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {     
            string message = string.Format("Thanks for click on a button.\r\nHostApp is {0}.{1} Version:{2}",
                                                _application.ComponentName, _application.Name, _application.Version);
            MessageBox.Show(message, _addinName, MessageBoxButtons.OK, MessageBoxIcon.Information);
      
            Ctrl.Dispose();
        }
   
        /// <summary>
        /// calls specific create method for office application type
        /// note: all applications has the same code except outlook (active inspector)
        /// </summary>
        public void CreateClassicUI()
        {

            if (_application.Application is Excel.Application)
            {
                CreateExcelUI();
            }
            else if (_application.Application is Word.Application)
            {
                CreateWordUI();
            }
            else if (_application.Application is Outlook.Application)
            {
                CreateOutlookUI();
            }
            else if (_application.Application is PowerPoint.Application)
            {
                CreatePowerPointUI();
            }
            else if (_application.Application is Access.Application)
            {
                CreateAccessUI();
            }
        }

        private void CreateExcelUI()
        {
            Excel.Application excelApp = _application.Application as Excel.Application;
            Office.CommandBar commandBar = excelApp.CommandBars.Add(_addinName + "Commandbar", Office.Enums.MsoBarPosition.msoBarTop, false, true);
            commandBar.Visible = true;

            // add a button to the toolbar
            Office.CommandBarButton commandBarBtn = (Office.CommandBarButton)commandBar.Controls.Add(Office.Enums.MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, true);
            commandBarBtn.Style = Office.Enums.MsoButtonStyle.msoButtonIconAndCaption;
            commandBarBtn.Caption = "ExcelButton";
            commandBarBtn.FaceId = 2;
            commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_ClickEvent);

            excelApp.DisposeChildInstances(false);
        }

        private void CreateOutlookUI()
        {
            Outlook.Application outlookApp = _application.Application as Outlook.Application;
            Office.CommandBar commandBar = outlookApp.ActiveExplorer().CommandBars.Add(_addinName + "CommandBar", Office.Enums.MsoBarPosition.msoBarTop, false, true);
            commandBar.Visible = true;

            // add a button to the toolbar
            Office.CommandBarButton commandBarBtn = (Office.CommandBarButton)commandBar.Controls.Add(Office.Enums.MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, true);
            commandBarBtn.Style = Office.Enums.MsoButtonStyle.msoButtonIconAndCaption;
            commandBarBtn.Caption = "OutlookButton";
            commandBarBtn.FaceId = 2;
            commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_ClickEvent);

            outlookApp.DisposeChildInstances(false);
        }

        private void CreateWordUI()
        {
            Word.Application wordApp = _application.Application as Word.Application;
            Office.CommandBar commandBar = wordApp.CommandBars.Add(_addinName + "Commandbar", Office.Enums.MsoBarPosition.msoBarTop, false, true);
            commandBar.Visible = true;

            // add a button to the toolbar
            Office.CommandBarButton commandBarBtn = (Office.CommandBarButton)commandBar.Controls.Add(Office.Enums.MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, true);
            commandBarBtn.Style = Office.Enums.MsoButtonStyle.msoButtonIconAndCaption;
            commandBarBtn.Caption = "WordButton";
            commandBarBtn.FaceId = 2;
            commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_ClickEvent);

            wordApp.DisposeChildInstances(false);
        }

        private void CreatePowerPointUI()
        {
            PowerPoint.Application powerApp = _application.Application as PowerPoint.Application;
            Office.CommandBar commandBar = powerApp.CommandBars.Add(_addinName + "CommandBar", Office.Enums.MsoBarPosition.msoBarTop, false, true);
            commandBar.Visible = true;

            // add a button to the toolbar
            Office.CommandBarButton commandBarBtn = (Office.CommandBarButton)commandBar.Controls.Add(Office.Enums.MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, true);
            commandBarBtn.Style = Office.Enums.MsoButtonStyle.msoButtonIconAndCaption;
            commandBarBtn.Caption = "PowerButton";
            commandBarBtn.FaceId = 2;
            commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_ClickEvent);

            powerApp.DisposeChildInstances(false);
        }

        private void CreateAccessUI()
        {
            Access.Application accessApp = _application.Application as Access.Application;
            Office.CommandBar commandBar = accessApp.CommandBars.Add(_addinName + "CommandBar", Office.Enums.MsoBarPosition.msoBarTop, false, true);
            commandBar.Visible = true;

            // add a button to the toolbar
            Office.CommandBarButton commandBarBtn = (Office.CommandBarButton)commandBar.Controls.Add(Office.Enums.MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, true);
            commandBarBtn.Style = Office.Enums.MsoButtonStyle.msoButtonIconAndCaption;
            commandBarBtn.Caption = "AccessButton";
            commandBarBtn.FaceId = 2;
            commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_ClickEvent); ;

            accessApp.DisposeChildInstances(false);
        }

        #endregion

        #region private statice Helper 
        
        private static string ReadTextFileFromRessource(string fileName)
        {
            fileName =  _addinName + "." + fileName;

            System.IO.Stream ressourceStream;
            System.IO.StreamReader textStreamReader;

            ressourceStream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(fileName);
            if (ressourceStream == null)
                throw (new System.IO.IOException("Error accessing resource Stream."));

            textStreamReader = new System.IO.StreamReader(ressourceStream);
            if (textStreamReader == null)
                throw (new System.IO.IOException("Error accessing resource File."));

            string text = textStreamReader.ReadToEnd();
            ressourceStream.Close();
            textStreamReader.Close();
            return text;
        }

        #endregion
    }
}
