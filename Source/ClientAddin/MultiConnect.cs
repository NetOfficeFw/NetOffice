using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NetOffice.Tools;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Tools;

namespace ClientAddin
{
    //[COMAddin("Compile Test Multi Addin", "", 3), ProgId("ClientAddin.MultiConnect"), Guid("C97F080F-5A1D-49F7-AE9B-1E53E6AC87EE")]
    //[RegistryLocation(RegistrySaveLocation.InstallScope), Programmable, Codebase, Lockback]
    //[MultiRegister(RegisterIn.Excel, RegisterIn.Word, RegisterIn.PowerPoint, RegisterIn.Access)]
    //public class MultiConnect : Office.Tools.COMAddin
    //{
    //    [RegisterErrorHandler]
    //    public static bool RegisterError(RegisterErrorMethodKind methodKind, Exception exception)
    //    {
    //        Console.WriteLine("MultiConnect::CreateExport");
    //        MessageDialog.ShowRegisterError(methodKind, exception, "Multi");
    //        return true;
    //    }

    //    [RegisterFunction(RegisterMode.CallAfter)]
    //    public static void Register(InstallScope scope)
    //    {
    //        MessageDialog.ShowRegister(scope, "Multi");
    //    }

    //    [UnRegisterFunction(RegisterMode.CallAfter)]
    //    public static void UnRegister(InstallScope scope)
    //    {
    //        MessageDialog.ShowUnRegister(scope, "Multi");
    //    }

    //    [RegExportFunction]
    //    private static RegExport CreateExport()
    //    {
    //        // Wird 4x aufgerufen ???
    //        Console.WriteLine("MultiConnect::CreateExport");
    //        return null;
    //    }
    //}
}
