using System;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Tools;
using Office = NetOffice.OfficeApi;
using NetOffice.ExcelApi.Tools;
using NetOffice.Tools.Native.Bridge;

namespace Excel06AddinCS4
{
    [COMAddin("Excel06 Sample Addin CS4", "Shim Addin Example", LoadBehavior.LoadAtStartup)]
    [ProgId("Excel06AddinCS4.Connect"), Guid("CF1DAC71-0E0F-411C-AA85-421AF9B35536"), Codebase, Timestamp]
    public class Addin : COMAddin
    {
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate void SayHello(string caption);

        [RegisterFunction(RegisterMode.CallAfter)]
        private static void Register(Type type, RegisterCall registerCall, InstallScope scope, OfficeRegisterKeyState keyState)
        {
            try
            {
                using (CdeclHandle libray = CdeclHandle.LoadLibrary(typeof(Addin), "Excel06AddinCS4.Shim.dll"))
                {
                    SayHello hello = libray.GetDelegateForFunctionPointer("SayHelloToTheWorld", typeof(SayHello)) as SayHello;
                    hello("Excel06AddinCS4");
                }
            }
            catch (Exception ex)
            {
                Office.Tools.Contribution.DialogUtils.ShowMessageBox(ex.ToString());
            }           
        }

        [RegisterErrorHandler]
        private static void RegisterError(RegisterErrorMethodKind methodKind, Exception exception)
        {
            Office.Tools.Contribution.DialogUtils.ShowRegisterError("Excel06AddinCS4", methodKind, exception);
        }
    }
}
