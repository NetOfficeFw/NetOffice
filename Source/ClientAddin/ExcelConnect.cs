using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NetOffice.Tools;
using Excel = NetOffice.ExcelApi;

namespace ClientAddin
{
    [COMAddin("Compile Test Excel Addin", "", 3), ProgId("ClientAddin.ExcelConnect"), Guid("F42C9AF1-E8B5-4480-ACF3-EBF097D914E4")]
    [RegistryLocation(RegistrySaveLocation.InstallScope), Programmable, Codebase, Lockback]
    public class ExcelConnect : Excel.Tools.COMAddin
    {
        public ExcelConnect()
        {     
            OnStartupComplete += ExcelConnect_OnStartupComplete;
        }

        private void ExcelConnect_OnStartupComplete(ref Array custom)
        {
            Utils.Dialog.ShowDiagnostics(false);
        }

        protected override void OnError(ErrorMethodKind methodKind, Exception exception)
        {
            Utils.Dialog.ShowError(exception, "An error occured");
        }

        [RegisterErrorHandler]
        private static void RegisterError(RegisterErrorMethodKind methodKind, Exception exception)
        {
            NetOffice.Settings.Default.MessageFilter.Enabled = true;           
            MessageDialog.ShowRegisterError(methodKind, exception, "Excel");
        }

        [RegisterFunction(RegisterMode.CallAfter)]
        private static void Register(Type type, RegisterCall registerCall, InstallScope scope, OfficeRegisterKeyState keyState)
        {
            MessageDialog.ShowRegister(type, registerCall, scope, keyState, "Excel");
        }

        [UnRegisterFunction(RegisterMode.CallAfter)]
        private static void UnRegister(Type type, RegisterCall registerCall, InstallScope scope, OfficeUnRegisterKeyState keyState)
        {
            MessageDialog.ShowUnRegister(type, registerCall, scope, keyState, "Excel");
        }
         
        [RegExportFunction]
        private static RegExport CreateExport(InstallScope scope, OfficeRegisterKeyState keyState)
        {
            Console.WriteLine("ExcelConnect::CreateExport");
            RegExport export = new RegExport();

            export.Add("Software\\Microsoft\\Office\\Excel\\Addins\\ClientAddin.ExcelConnect",
                new RegExportValue[] {
                                       new RegExportValue("Value 1", Microsoft.Win32.RegistryValueKind.DWord, 144),
                                       new RegExportValue("Value 2", Microsoft.Win32.RegistryValueKind.QWord, 1433689),
                                       new RegExportValue("Value 3", Microsoft.Win32.RegistryValueKind.String,"Hello Excel"),
                                       new RegExportValue("Value 4", Microsoft.Win32.RegistryValueKind.ExpandString, "Hello World"),
                                       new RegExportValue("Value 5", Microsoft.Win32.RegistryValueKind.MultiString, new string[] { "Item 1","Item 2","Item 3"}),
                                       new RegExportValue("Value 6", Microsoft.Win32.RegistryValueKind.Binary,new byte[] { 1, 2, 3, 4, 0, 7, 15, 31, 127, 255}),
                                    });

            IList<RegExportValue> settings = export.Add("Software\\Microsoft\\Office\\Excel\\Addins\\ClientAddin.ExcelConnect\\Settings");
            settings.Add(new RegExportValue("", "I was here"));
            settings.Add(new RegExportValue("Setting 1", "Enabled"));
            settings.Add(new RegExportValue("Setting 2", "Disabled"));
            
            return export;
        }
    } 
}
