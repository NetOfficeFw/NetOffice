using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Tools;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using NetOffice.WordApi.Tools;
/*
   Register Addin Example
*/
namespace Excel04AddinCS4
{
    [COMAddin("Word04 Sample Addin CS4", "Register Addin Example", LoadBehavior.LoadAtStartup)]
    [ProgId("Word04AddinCS4.Connect"), Guid("B13E5DB8-1F31-45DB-ABFA-C08D126B5898"), Codebase, Timestamp]
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
            Office.Tools.Contribution.DialogUtils.ShowRegisterError("Word04AddinCS4", methodKind, exception);
        }
    }
}