using System;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

using Excel = NetOffice.ExcelApi;
using Office = NetOffice.OfficeApi;

using NetOffice.ExcelApi.Enums;
using NetOffice.OfficeApi.Enums;

namespace ExcelAddinCSharp
{
    [ComVisible(true)]
    [GuidAttribute("EF294277-3DF1-425F-AC90-8687DEAC280D"), ProgId("ExcelAddinCSharp.TestAddin")]
    public class TestAddin : IDTExtensibility2, IRibbonExtensibility, ICustomTaskPaneConsumer
    {
        private static readonly string _addinRegistryKey = "Software\\Microsoft\\Office\\Excel\\AddIns\\";
        private static readonly string _prodId = "ExcelAddinCSharp.TestAddin";
        private static readonly string _addinName = "TestAddin C# Excel";
        
        Excel.Application _excelApplication;

        private bool _ribbonUIPassed;
        private bool _taskPanePassed;

        public bool RibbonUI
        {
            get
            {
                return _ribbonUIPassed;
            }
        }

        public bool TaskPanePassed
        {
            get
            {
                return _taskPanePassed;
            }
        }

        #region COM Register Functions

        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type type)
        {
            try
            {
                // add codebase value
                Assembly thisAssembly = Assembly.GetAssembly(typeof(TestAddin));
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
                rk.SetValue("Description", "TestAddin C# Excel");
                rk.Close();
            }
            catch (Exception ex)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", ex.Message, Environment.NewLine);
                //MessageBox.Show("An error occured." + details, "Register " + _addinName, MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                //MessageBox.Show("An error occured." + details, "Unregister " + _addinName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region ICustomTaskPaneConsumer Member

        private Office.ICTPFactory _myCtpFactory;
        private Office._CustomTaskPane _myPane;
        private SampleControl _myControl;

        public void CTPFactoryAvailable(object CTPFactoryInst)
        {
            _myCtpFactory = new NetOffice.OfficeApi.ICTPFactory(_excelApplication, CTPFactoryInst);
            _myPane = _myCtpFactory.CreateCTP("ExcelAddinCSharp.SampleControl", "NetOffice Sample Task Pane", Type.Missing);
            _myPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            _myPane.Visible = true;
            _myControl = _myPane.ContentControl as SampleControl;
            _taskPanePassed = true;
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

                _excelApplication = new Excel.Application(null, Application);

                Office.COMAddIn addin = new Office.COMAddIn(null, AddInInst);
                addin.Object = this;
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
                if (null != _excelApplication)
                    _excelApplication.Dispose();
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
                _ribbonUIPassed = true;
                return ReadString("RibbonUI.xml");
            }
            catch (Exception throwedException)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine);
                //MessageBox.Show("An error occured in GetCustomUI." + details, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                //MessageBox.Show("An error occured in OnAction." + details, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            fileName = "ExcelAddinCSharp." + fileName;

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
