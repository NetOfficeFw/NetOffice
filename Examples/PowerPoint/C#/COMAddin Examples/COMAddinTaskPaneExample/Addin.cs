using System;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using Extensibility;

using NetOffice;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using PowerPoint = NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Enums;

namespace COMAddinTaskPaneExampleCS4
{
    [GuidAttribute("40C8F944-BD8D-41D7-8314-2C99343F148C"), ProgId("PPointAddinCS4.TaskPaneAddin"), ComVisible(true)]
    public class Addin : IDTExtensibility2, Office.ICustomTaskPaneConsumer
    {
        private static readonly string _addinOfficeRegistryKey  = "Software\\Microsoft\\Office\\PowerPoint\\AddIns\\";
        private static readonly string _progId                  = "PPointAddinCS4.TaskPaneAddin";
        private static readonly string _addinFriendlyName       = "NetOffice Sample Addin in C#";
        private static readonly string _addinDescription        = "NetOffice Sample Addin with custom Task Pane";

        private static SampleControl _sampleControl;
        private static PowerPoint.Application _powerApplication;

        internal static PowerPoint.Application Application { get { return _powerApplication; } }

        #region ICustomTaskPaneConsumer Member

        public void CTPFactoryAvailable(object CTPFactoryInst)
        {
            try
            {
                Office.ICTPFactory ctpFactory = new NetOffice.OfficeApi.ICTPFactory(_powerApplication, CTPFactoryInst);
                Office._CustomTaskPane taskPane = ctpFactory.CreateCTP(typeof(Addin).Assembly.GetName().Name + ".SampleControl", "NetOffice Sample Pane(CS4)", Type.Missing);
                taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                taskPane.Width = 300;
                taskPane.Visible = true;
                _sampleControl = taskPane.ContentControl as SampleControl;
                ctpFactory.Dispose();
            }
            catch (Exception exception)
            {
                string message = string.Format("An error occured.{0}{0}{1}", Environment.NewLine, exception.Message);
                MessageBox.Show(message, _progId, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region IDTExtensibility2 Members

        void IDTExtensibility2.OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                _powerApplication = new PowerPoint.Application(null, Application);
            }
            catch (Exception exception)
            {
                string message = string.Format("An error occured.{0}{0}{1}", Environment.NewLine, exception.Message);
                MessageBox.Show(message, _progId, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        void IDTExtensibility2.OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                if (null != _powerApplication)
                    _powerApplication.Dispose();
            }
            catch (Exception exception)
            {
                string message = string.Format("An error occured.{0}{0}{1}", Environment.NewLine, exception.Message);
                MessageBox.Show(message, _progId, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        void IDTExtensibility2.OnStartupComplete(ref Array custom)
        {

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

                // register addin in powerpoint
                Registry.CurrentUser.CreateSubKey(_addinOfficeRegistryKey + _progId);
                RegistryKey regKeyExcel = Registry.CurrentUser.OpenSubKey(_addinOfficeRegistryKey + _progId, true);
                regKeyExcel.SetValue("LoadBehavior", Convert.ToInt32(3));
                regKeyExcel.SetValue("FriendlyName", _addinFriendlyName);
                regKeyExcel.SetValue("Description", _addinDescription);
                regKeyExcel.Close();
            }
            catch (Exception ex)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", ex.Message, Environment.NewLine);
                MessageBox.Show("An error occured." + details, "Register " + _progId, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type type)
        {
            try
            {
                // unregister addin
                Registry.ClassesRoot.DeleteSubKey(@"CLSID\{" + type.GUID.ToString().ToUpper() + @"}\Programmable", false);

                // unregister addin in office
                Registry.CurrentUser.DeleteSubKey(_addinOfficeRegistryKey + _progId, false);

            }
            catch (Exception throwedException)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine);
                MessageBox.Show("An error occured." + details, "Unregister " + _progId, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion
    }
}
