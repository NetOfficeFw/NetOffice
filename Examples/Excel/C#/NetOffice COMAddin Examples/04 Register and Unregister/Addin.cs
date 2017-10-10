using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Tools;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using NetOffice.ExcelApi.Tools;
/*
   Register Addin Example
*/
namespace Excel04AddinCS4
{
    [COMAddin("Excel04 Sample Addin CS4", "Register Addin Example", LoadBehavior.LoadAtStartup)]
    [ProgId("Excel04AddinCS4.Connect"), Guid("F0813E2B-C963-431E-A0B3-971A1F70E8A8"), Codebase, Timestamp]
    [RegistryLocation(RegistrySaveLocation.InstallScopeCurrentUser)]
    public class Addin : COMAddin
    {       
        [RegisterFunction(RegisterMode.CallAfter)]  // We want that NetOffice call this method after register
        private static void Register(Type type, RegisterCall registerCall, InstallScope scope, OfficeRegisterKeyState keyState)
        {

        }
        
        [UnRegisterFunction(RegisterMode.CallBeforeAndAfter)] // We want that NetOffice call this method before and after unregister
        private static void UnRegister(Type type, RegisterCall registerCall, InstallScope scope, OfficeUnRegisterKeyState keyState)
        {

        }

        // An unexpected error occured in register or unregister action
        [RegisterErrorHandler]
        private static void RegisterError(RegisterErrorMethodKind methodKind, Exception exception)
        {
            Office.Tools.Contribution.DialogUtils.ShowRegisterError("Excel04AddinCS4", methodKind, exception);
        }
    }
}