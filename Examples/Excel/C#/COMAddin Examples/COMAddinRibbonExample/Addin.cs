using System;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using Extensibility;

using LateBindingApi.Core;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;

namespace COMAddinRibbonExample
{
    [GuidAttribute("182f0e9a-07db-4375-8372-77e1637d4891"), ProgId("COMAddinRibbonExampleCSharp.Addin"), ComVisible(true)]
    public class Addin : IDTExtensibility2, Office.IRibbonExtensibility
    {
        private static readonly string _addinOfficeRegistryKey  = "Software\\Microsoft\\Office\\Excel\\AddIns\\";
        private static readonly string _prodId                  = "COMAddinRibbonExampleCSharp.Addin";
        private static readonly string _addinFriendlyName       = "NetOffice Sample Addin in C#";
        private static readonly string _addinDescription        = "NetOffice Sample Addin with custom Ribbon UI";

        Excel.Application _excelApplication;

        #region IDTExtensibility2 Members

        void IDTExtensibility2.OnStartupComplete(ref Array custom)
        {

        }

        void IDTExtensibility2.OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                // Initialize NetOffice
                LateBindingApi.Core.Factory.Initialize();

                _excelApplication = new Excel.Application(null, Application);
            }
            catch (Exception exception)
            {
                string message = string.Format("An error occured.{0}{0}{1}", Environment.NewLine, exception.Message);
                MessageBox.Show(message, _addinFriendlyName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        void IDTExtensibility2.OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                _excelApplication.Dispose();
            }
            catch (Exception exception)
            {
                string message = string.Format("An error occured.{0}{0}{1}", Environment.NewLine, exception.Message);
                MessageBox.Show(message, _addinFriendlyName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        void IDTExtensibility2.OnAddInsUpdate(ref Array custom)
        {

        }

        void IDTExtensibility2.OnBeginShutdown(ref Array custom)
        {

        }

        #endregion

        #region IRibbonExtensibility Members

        public string GetCustomUI(string RibbonID)
        {
            try
            {
                return ReadRessourceFile("RibbonUI.xml");
            }
            catch (Exception throwedException)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine);
                MessageBox.Show("An error occured in GetCustomUI: " + details, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return "";
            }
        }

        public void OnAction(Office.IRibbonControl control)
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
                MessageBox.Show("An error occured in OnAction: " + details, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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

                // register addin in Excel
                Registry.CurrentUser.CreateSubKey(_addinOfficeRegistryKey + _prodId);
                RegistryKey regKeyExcel = Registry.CurrentUser.OpenSubKey(_addinOfficeRegistryKey + _prodId, true);
                regKeyExcel.SetValue("LoadBehavior", Convert.ToInt32(3));
                regKeyExcel.SetValue("FriendlyName", _addinFriendlyName);
                regKeyExcel.SetValue("Description", _addinDescription);
                regKeyExcel.Close();
            }
            catch (Exception ex)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", ex.Message, Environment.NewLine);
                MessageBox.Show("An error occured." + details, "Register Addin", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                Registry.CurrentUser.DeleteSubKey(_addinOfficeRegistryKey + _prodId);

            }
            catch (Exception throwedException)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine);
                MessageBox.Show("An error occured." + details, "Unregister Addin", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Private Helper Methods

        /// <summary>
        /// reads text file from ressource
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        protected internal static string ReadRessourceFile(string fileName)
        {
            Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            System.IO.Stream ressourceStream = assembly.GetManifestResourceStream(assembly.GetName().Name + "." + fileName);
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
