using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NetOffice.Tools;
using Word = NetOffice.WordApi;

namespace ClientAddin
{
    //[COMAddin("Compile Test Word Addin", "", 3), ProgId("ClientAddin.WordConnect"), Guid("D2B4EA8A-AD49-4ECE-83A3-FEC1819BC5AE")]
    //[Programmable, Codebase, Lockback]
    //public class WordConnect : Word.Tools.COMAddin
    //{
    //    [RegisterErrorHandler]
    //    public static bool RegisterError(RegisterErrorMethodKind methodKind, Exception exception)
    //    {
    //        MessageDialog.ShowRegisterError(methodKind, exception, "Word");
    //        return true;
    //    }

    //    [RegisterFunction(RegisterMode.CallAfter)]
    //    public static void Register(Type type, RegisterCall registerCall)
    //    {
    //        MessageDialog.ShowRegister(type, registerCall, "Word");
    //    }

    //    [UnRegisterFunction(RegisterMode.CallAfter)]
    //    public static void UnRegister(Type type, RegisterCall registerCall)
    //    {
    //        MessageDialog.ShowUnRegister(type, registerCall, "Word");
    //    }

    //    [RegExportFunction]
    //    private static RegExport CreateExport(InstallScope scope)
    //    {
    //        Console.WriteLine("WordConnect::CreateExport");

    //        RegExport export = new RegExport();

    //        export.Add("Software\\MyCompany",
    //            new RegExportValue[] { new RegExportValue("Value 1", "Hello World"),
    //                                   new RegExportValue("Value 2", "Hello Word"),
    //                                   new RegExportValue("Value 3", Microsoft.Win32.RegistryValueKind.DWord, 3) });

    //        return export;
    //    }
    //}
}
