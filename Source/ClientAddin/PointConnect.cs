using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NetOffice.Tools;
using Point = NetOffice.PowerPointApi;

namespace ClientAddin
{
    [COMAddin("Compile Test Point Addin", "", 3), ProgId("ClientAddin.PointConnect"), Guid("A8767F92-12DC-47BC-A8D0-FA42FB99377C")]
    [Programmable, Codebase, Lockback]
    public class PointConnect : Point.Tools.COMAddin
    {
        public PointConnect()
        {
            OnStartupComplete += PointConnect_OnStartupComplete;
        }

        private void PointConnect_OnStartupComplete(ref Array custom)
        {
            var hwnd = Utils.Application.HWND;
            Utils.Dialog.ShowMessageBox(hwnd.ToString(), DialogResult.OK);              
        }

        [RegisterErrorHandler]
        public static bool RegisterError(RegisterErrorMethodKind methodKind, Exception exception)
        {
            MessageDialog.ShowRegisterError(methodKind, exception, "Point");
            return true;
        }

        [RegisterFunction(RegisterMode.CallAfter)]
        public static void Register(InstallScope scope)
        {
            MessageDialog.ShowRegister(scope, "Point");
        }

        [UnRegisterFunction(RegisterMode.CallAfter)]
        public static void UnRegister(InstallScope scope)
        {
            MessageDialog.ShowUnRegister(scope, "Point");
        }

        [RegExportFunction]
        private static RegExport CreateExport()
        {
            Console.WriteLine("PointConnect::CreateExport");
            return null;
        }
     }
}
