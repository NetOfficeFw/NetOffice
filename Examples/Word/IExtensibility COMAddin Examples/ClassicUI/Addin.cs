using System;
using Extensibility;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Word = NetOffice.WordApi;
using Office = NetOffice.OfficeApi;
using NetOffice.WordApi.Enums;
using NetOffice.OfficeApi.Enums;

namespace COMAddinClassicExampleCS4
{
    [Guid("1C401FCE-7D5E-4C6D-AE88-71BF01FC159B"), ProgId("WordAddinCS4.SimpleAddin"), ComVisible(true)]
    public class Addin : IDTExtensibility2
    {
        private static readonly string _addinOfficeRegistryKey  = "Software\\Microsoft\\Office\\Word\\AddIns\\";
        private static readonly string _prodId                  = "WordAddinCS4.SimpleAddin";
        private static readonly string _addinFriendlyName       = "NetOffice Sample Addin in C#";
        private static readonly string _addinDescription        = "NetOffice Sample Addin with custom classic UI";

        // gui elements
        private static readonly string _toolbarName             = "Sample Toolbar CS4";
        private static readonly string _toolbarButtonName       = "Sample ToolbarButton CS4";
        private static readonly string _toolbarPopupName        = "Sample ToolbarPopup CS4";
        private static readonly string _menuName                = "Sample Menu CS4";
        private static readonly string _menuButtonName          = "Sample Button CS4";
        private static readonly string _contextName             = "Sample ContextMenu CS4";
        private static readonly string _contextMenuButtonName   = "Sample ContextButton CS4";

        private Word.Application _wordApplication;
        private Word.Template    _normalDotTemplate; 

        #region IDTExtensibility2 Members

        void IDTExtensibility2.OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                _wordApplication = new Word.Application(null, Application);
            }
            catch (Exception exception)
            {
                string message = string.Format("An error occured.{0}{0}{1}", Environment.NewLine, exception.Message);
                MessageBox.Show(message, _prodId, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        void IDTExtensibility2.OnStartupComplete(ref Array custom)
        {
            try
            {
                GetNormalDotTemplate();
                RemoveGui();
                SetupGui();
            }
            catch (Exception exception)
            {
                string message = string.Format("An error occured.{0}{0}{1}", Environment.NewLine, exception.Message);
                MessageBox.Show(message, _prodId, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        void IDTExtensibility2.OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                if (null != _wordApplication)
                {
                    // word ignores the temporary parameter in created menus(not toolbars) and save menu settings to normal.dot 
                    RemoveGui();
                    _wordApplication.Dispose();
                }
            }
            catch (Exception exception)
            {
                string message = string.Format("An error occured.{0}{0}{1}", Environment.NewLine, exception.Message);
                MessageBox.Show(message, _prodId, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
  
        void IDTExtensibility2.OnAddInsUpdate(ref Array custom)
        {
        }

        void IDTExtensibility2.OnBeginShutdown(ref Array custom)
        {
        }

        #endregion

        #region COM Register Functions

        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type type)
        {
            try
            {
                // add codebase value
                Assembly thisAssembly = Assembly.GetAssembly(typeof(Addin));
                RegistryKey key = Registry.ClassesRoot.CreateSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\InprocServer32\\1.0.0.0");
                key.SetValue("CodeBase", thisAssembly.CodeBase);
                key.Close();

                Registry.ClassesRoot.CreateSubKey(@"CLSID\{" + type.GUID.ToString().ToUpper() + @"}\Programmable");

                // add bypass key
                // http://support.microsoft.com/kb/948461
                key = Registry.ClassesRoot.CreateSubKey("Interface\\{000C0601-0000-0000-C000-000000000046}");
                string defaultValue = key.GetValue("") as string;
                if (null == defaultValue)
                    key.SetValue("", "Office .NET Framework Lockback Bypass Key");
                key.Close();

                // add word addin key
                Registry.CurrentUser.CreateSubKey(_addinOfficeRegistryKey + _prodId);
                RegistryKey rk = Registry.CurrentUser.OpenSubKey(_addinOfficeRegistryKey + _prodId, true);
                rk.SetValue("LoadBehavior", Convert.ToInt32(3));
                rk.SetValue("FriendlyName", _addinFriendlyName);
                rk.SetValue("Description", _addinDescription);
                rk.Close();
            }
            catch (Exception ex)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", ex.Message, Environment.NewLine);
                MessageBox.Show("An error occured." + details, "Register " + _prodId, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type type)
        {
            try
            {
                Registry.ClassesRoot.DeleteSubKey(@"CLSID\{" + type.GUID.ToString().ToUpper() + @"}\Programmable", false);
                Registry.CurrentUser.DeleteSubKey(_addinOfficeRegistryKey + _prodId, false);
            }
            catch (Exception throwedException)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine);
                MessageBox.Show("An error occured." + details, "Unregister " + _prodId, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Private Helper Methods
        
        /// <summary>
        /// returns normal.dot template
        /// </summary>
        private void GetNormalDotTemplate()
        {
            foreach (Word.Template installedTemplate in _wordApplication.Templates)
            {
                if(installedTemplate.Name.StartsWith("normal", StringComparison.InvariantCultureIgnoreCase))
                {
                    _normalDotTemplate = installedTemplate;
                    return;
                }
                installedTemplate.Dispose();
            }
        }
        
        #endregion

        #region Create/Remove UI

        /// <summary>
        /// removes gui elements if exists
        /// </summary>
        private void RemoveGui()
        {
            _wordApplication.CustomizationContext = _normalDotTemplate;

            Office.CommandBar menuBar = _wordApplication.CommandBars["Menu Bar"];
            Office.CommandBar contextBar = _wordApplication.CommandBars["Text"];

            Office.CommandBarControl control = menuBar.FindControl(System.Type.Missing, System.Type.Missing, _menuName, System.Type.Missing, false);
            if (null != control)
                control.Delete();

            control = contextBar.FindControl(System.Type.Missing, System.Type.Missing, _contextName, System.Type.Missing, false);
            if (null != control)
                control.Delete();

            menuBar.Dispose();
            contextBar.Dispose();
        }

        /// <summary>
        /// creates gui elements
        /// </summary>
        private void SetupGui()
        {
            /*
            // How to: Add Commands to Shortcut Menus in Word
            // http://msdn.microsoft.com/de-de/library/dd554969.aspx             
            */

            _wordApplication.CustomizationContext = _normalDotTemplate;

            /* create commandbar */
            Office.CommandBar commandBar = _wordApplication.CommandBars.Add(_toolbarName, MsoBarPosition.msoBarTop, System.Type.Missing, true);            
            commandBar.Visible = true;
        
            // add popup to commandbar
            Office.CommandBarPopup commandBarPop = (Office.CommandBarPopup)commandBar.Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
            commandBarPop.Caption = _toolbarPopupName;
            commandBarPop.Tag = _toolbarPopupName;

            // add a button to the popup
            Office.CommandBarButton  commandBarBtn = (Office.CommandBarButton)commandBarPop.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
            commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption;
            commandBarBtn.FaceId = 9;
            commandBarBtn.Caption = _toolbarButtonName;
            commandBarBtn.Tag = _toolbarButtonName;
            commandBarBtn.ClickEvent += new NetOffice.OfficeApi.CommandBarButton_ClickEventHandler(commandBarBtn_ClickEvent);

            /* create menu */
            commandBar = _wordApplication.CommandBars["Menu Bar"];
            
            // add popup to menu bar
            commandBarPop = (Office.CommandBarPopup)commandBar.Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
            commandBarPop.Caption = _menuName;
            commandBarPop.Tag = _menuName;

            // add a button to the popup
            commandBarBtn = (Office.CommandBarButton)commandBarPop.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
            commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption;
            commandBarBtn.FaceId = 9;
            commandBarBtn.Caption = _menuButtonName;
            commandBarBtn.Tag = _menuButtonName;
            commandBarBtn.ClickEvent += new NetOffice.OfficeApi.CommandBarButton_ClickEventHandler(commandBarBtn_ClickEvent);

            /* create context menu */ 
            commandBarPop = (Office.CommandBarPopup)_wordApplication.CommandBars["Text"].Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
            commandBarPop.Caption = _contextName;
            commandBarPop.Tag     = _contextName;

            // add a button to the popup
            commandBarBtn = (Office.CommandBarButton)commandBarPop.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
            commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption;
            commandBarBtn.Caption = _contextMenuButtonName;
            commandBarBtn.Tag = _contextMenuButtonName;
            commandBarBtn.FaceId = 9;
            commandBarBtn.ClickEvent += new NetOffice.OfficeApi.CommandBarButton_ClickEventHandler(commandBarBtn_ClickEvent);

            _normalDotTemplate.Saved = true;
        }

        #endregion

        #region UI Trigger

        /// <summary>
        /// Click event trigger from created buttons. incoming call comes from word application thread.
        /// </summary>
        /// <param name="Ctrl"></param>
        /// <param name="CancelDefault"></param>
        void commandBarBtn_ClickEvent(NetOffice.OfficeApi.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                string message = string.Format("Click from Button {0}.", Ctrl.Caption);
                MessageBox.Show(message, _prodId, MessageBoxButtons.OK, MessageBoxIcon.Information);
                Ctrl.Dispose();
            }
            catch (Exception exception)
            {
                string message = string.Format("An error occured.{0}{0}{1}", Environment.NewLine, exception.Message);
                MessageBox.Show(message, _prodId, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion
    }
}
