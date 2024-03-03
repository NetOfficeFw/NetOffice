using System;
using System.Reflection; 
using System.Windows.Forms;
using System.Collections.Generic;
using Microsoft.Win32;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using Extensibility;
using NetOffice;
using Office = NetOffice.OfficeApi;
using Excel = NetOffice.ExcelApi;
using Word = NetOffice.WordApi;
using Outlook = NetOffice.OutlookApi;
using PowerPoint = NetOffice.PowerPointApi;
using Access = NetOffice.AccessApi;

namespace Super2AddinCS4
{ 
    /// <summary>
    /// the addin class
    /// </summary>
    [Guid("39AA1CDE-F186-456B-ADF6-30D03352C131"), ProgId("Super2AddinCS4.Connect"), ComVisible(true)]
    public class Addin : IDTExtensibility2, Office.Native.IRibbonExtensibility
    {
        private static readonly string _progId                  = "Super2AddinCS4.Connect";
        private static readonly string _addinFriendlyName       = "NetOffice Sample Addin in C#";
        private static readonly string _addinDescription        = "NetOffice Sample Addin for multipe Office Applications";

        #region Fields 

        private ICOMObject _application;
        private string _hostApplicationName;

        #endregion

        #region IDTExtensibility2 Members
         
        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                _application = Core.Default.CreateObjectFromComProxy(null, Application, true);

                /*
                * _application is stored as ICOMObject the common base type for all reference types in NetOffice
                * because this addin is loaded in different office application.
                * 
                * with the CreateObjectFromComProxy method the _application instance is created as corresponding wrapper based on the com proxy type
                */

                if (_application is Excel.Application)
                   _hostApplicationName = "Excel";
                else if (_application is Word.Application)
                   _hostApplicationName = "Word";
                else if (_application is Outlook.Application)
                   _hostApplicationName = "Outlook";
                else if (_application is PowerPoint.Application)
                   _hostApplicationName = "PowerPoint";
                else if (_application is Access.Application)
                   _hostApplicationName = "Access";
            }
            catch (Exception exception)
            {
                if(_hostApplicationName != null)
                    OfficeRegistry.LogErrorMessage(_hostApplicationName, _progId, "Error occured in OnConnection. ", exception);
            }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                if (null != _application)
                    _application.Dispose();
            }
            catch (Exception exception)
            {
                OfficeRegistry.LogErrorMessage(_hostApplicationName, _progId, "Error occured in OnDisconnection. ", exception);
            }
        }

        public void OnStartupComplete(ref Array custom)
        {

        }

        public void OnAddInsUpdate(ref Array custom)
        {

        }

        public void OnBeginShutdown(ref Array custom)
        {

        }

        #endregion

        #region IRibbonExtensibility Members

        public string GetCustomUI(string RibbonID)
        {
            try
            {
                return ReadTextFileFromRessource("RibbonUI.xml");
            }
            catch (Exception exception)
            {
                OfficeRegistry.LogErrorMessage(_hostApplicationName, _progId, "Error occured in GetCustomUI. ", exception);
                return "";
            }
        }

        public void OnAction(Office.IRibbonControl control)
        {
            try
            {
                string message = string.Format("Thanks for click on a Ribbon.\r\nHostApp is {0}.{1} Version:{2}",
                                              TypeDescriptor.GetComponentName(_application.UnderlyingObject), _application.UnderlyingTypeName, Invoker.Default.PropertyGet(_application, "Version"));
                               
                MessageBox.Show(message, _progId, MessageBoxButtons.OK, MessageBoxIcon.Information);
                 
            }
            catch (Exception exception)
            {
                string message = string.Format("An error occured.{0}{0}{1}", Environment.NewLine, exception.Message);
                MessageBox.Show(message, _progId, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region COM Register Functions

        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type type)
        {
            try
            {
                // add codebase entry
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

                OfficeRegistry.CreateAddinKey("Excel" , _progId, _addinFriendlyName, _addinDescription);
                OfficeRegistry.CreateAddinKey("Word", _progId, _addinFriendlyName, _addinDescription);
                OfficeRegistry.CreateAddinKey("Outlook", _progId, _addinFriendlyName, _addinDescription);
                OfficeRegistry.CreateAddinKey("PowerPoint", _progId, _addinFriendlyName, _addinDescription);
                OfficeRegistry.CreateAddinKey("Access", _progId, _addinFriendlyName, _addinDescription);
            }
            catch (Exception exception)
            {
                string message = string.Format("An error occured.{0}{0}{1}", Environment.NewLine, exception.Message);
                MessageBox.Show(message, _progId, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type type)
        {
            try
            {
                Registry.ClassesRoot.DeleteSubKey(@"CLSID\{" + type.GUID.ToString().ToUpper() + @"}\Programmable", false);

                OfficeRegistry.DeleteAddinKey(OfficeRegistry.Excel + _progId);
                OfficeRegistry.DeleteAddinKey(OfficeRegistry.Word + _progId);
                OfficeRegistry.DeleteAddinKey(OfficeRegistry.Outlook + _progId);
                OfficeRegistry.DeleteAddinKey(OfficeRegistry.PowerPoint + _progId);
                OfficeRegistry.DeleteAddinKey(OfficeRegistry.Access + _progId);
            }
            catch (Exception exception)
            {
                string message = string.Format("An error occured.{0}{0}{1}", Environment.NewLine, exception.Message);
                MessageBox.Show(message, _progId, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Private Helper 
        
        private static string ReadTextFileFromRessource(string fileName)
        {
            Assembly assembly = typeof(Addin).Assembly;
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
