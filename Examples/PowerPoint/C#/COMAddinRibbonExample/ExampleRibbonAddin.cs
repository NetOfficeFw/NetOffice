using System;
using System.Reflection;
using System.Windows.Forms; 
using Microsoft.Win32;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

using PowerPoint = NetOffice.PowerPointApi;
using Office = NetOffice.OfficeApi;

using NetOffice.PowerPointApi.Enums;
using NetOffice.OfficeApi.Enums;

namespace COMAddinRibbonExample
{
    [ComVisible(true)]
    [GuidAttribute("F09B3141-ADBB-44c1-A6DC-A7F49498AB0F"), ProgId("WPowerPointibbonAddinCSharp.Addin")]
    public class ExampleRibbonAddin : IDTExtensibility2, IRibbonExtensibility
    {
        private static readonly string _addinRegistryKey = "Software\\Microsoft\\Office\\PowerPoint\\AddIns\\";
        private static readonly string _prodId           = "PowerPointRibbonAddinCSharp.Addin";
        private static readonly string _addinName        = "C# PowerPointRibbonAddin";

        PowerPoint.Application _powerApplication;

        #region COM Register Functions

        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type type)
        {
            try
            {
                // add codebase value
                Assembly thisAssembly = Assembly.GetAssembly(typeof(ExampleRibbonAddin));
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

                // add excel addin key
                Registry.ClassesRoot.CreateSubKey(@"CLSID\{" + type.GUID.ToString().ToUpper() + @"}\Programmable");
                Registry.CurrentUser.CreateSubKey(_addinRegistryKey + _prodId);
                RegistryKey rk = Registry.CurrentUser.OpenSubKey(_addinRegistryKey + _prodId, true);
                rk.SetValue("LoadBehavior", Convert.ToInt32(3));
                rk.SetValue("FriendlyName", _addinName);
                rk.SetValue("Description", "NetOffice COMAddinExample with ribbon UI");
                rk.Close();
            }
            catch (Exception ex)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", ex.Message, Environment.NewLine);
                MessageBox.Show("An error occured." + details, "Register " + _addinName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type type)
        {
            try
            {
                Registry.ClassesRoot.DeleteSubKey(@"CLSID\{" + type.GUID.ToString().ToUpper() + @"}\Programmable", false);
                Registry.CurrentUser.DeleteSubKey(_addinRegistryKey + _prodId);
            }
            catch (ArgumentException)
            {
                // key is already deleted
                ;
            }
            catch (Exception throwedException)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine);
                MessageBox.Show("An error occured." + details, "Unregister " + _addinName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region IDTExtensibility2 Members

        void IDTExtensibility2.OnAddInsUpdate(ref Array custom)
        {

        }

        void IDTExtensibility2.OnBeginShutdown(ref Array custom)
        {

        }

        void IDTExtensibility2.OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                // initialize api
                LateBindingApi.Core.Factory.Initialize();

                _powerApplication = new PowerPoint.Application(null, Application);
            }
            catch (Exception throwedException)
            {
                // dont show Dialogs or MessageBoxes in IDTExtensibility2 Functions
                // we save the error info in addin registry key

                string details = string.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine);

                RegistryKey rk = Registry.CurrentUser.OpenSubKey(_addinRegistryKey + _prodId, true);
                rk.SetValue("LastError", "An error occured in OnConnection.");
                rk.SetValue("LastException", throwedException.Message);
                rk.Close();
            }
        }

        void IDTExtensibility2.OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                if (null != _powerApplication)
                    _powerApplication.Dispose();
            }
            catch (Exception throwedException)
            {
                // dont show Dialogs or MessageBoxes in IDTExtensibility2 Functions
                // we save the error info in addin registry key

                string details = string.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine);

                RegistryKey rk = Registry.CurrentUser.OpenSubKey(_addinRegistryKey + _prodId, true);
                rk.SetValue("LastError", "An error occured in OnDisconnection.");
                rk.SetValue("LastException", throwedException.Message);
                rk.Close();
            }
        }

        void IDTExtensibility2.OnStartupComplete(ref Array custom)
        {

        }

        #endregion

        #region IRibbonExtensibility Members

        public string GetCustomUI(string RibbonID)
        {
            try
            {
                return ReadString("RibbonUI.xml");
            }
            catch (Exception throwedException)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine);
                MessageBox.Show("An error occured in GetCustomUI." + details, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
               return "";
            }
        }

        #endregion

        #region Ribbon Gui Trigger

        public void OnAction(IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "customButton1":
                        MessageBox.Show("This is the first sample button.");
                        break;
                    case "customButton2":
                        MessageBox.Show("This is the second sample button.");
                        break;
                    default:
                        MessageBox.Show("Unkown Control Id: " + control.Id);
                        break;
                }
            }
            catch (Exception throwedException)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine);
                MessageBox.Show("An error occured." + details, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Private Helper

        /// <summary>
        /// reads text from ressource
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private static string ReadString(string fileName)
        {
            fileName = "COMAddinRibbonExample." + fileName;

            System.IO.Stream ressourceStream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(fileName);
            if (ressourceStream == null)
                throw (new System.IO.IOException("Error accessing resource Stream."));

            System.IO.StreamReader textStreamReader = new System.IO.StreamReader(ressourceStream);
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
