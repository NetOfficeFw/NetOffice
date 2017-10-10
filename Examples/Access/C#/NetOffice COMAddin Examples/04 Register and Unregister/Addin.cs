using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Tools;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using Access = NetOffice.AccessApi;
using NetOffice.AccessApi.Enums;
using NetOffice.AccessApi.Tools;
/*
   Register Addin Example
*/
namespace Access04AddinCS4
{
    [COMAddin("Access04 Sample Addin CS4", "Register Addin Example", LoadBehavior.LoadAtStartup)]
    [ProgId("Access04AddinCS4.Connect"), Guid("CBDE34D8-060F-45EC-89C1-6D6687774ED5"), Codebase, Timestamp]
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
            Office.Tools.Contribution.DialogUtils.ShowRegisterError("Access04AddinCS4", methodKind, exception);
        }
    }
}