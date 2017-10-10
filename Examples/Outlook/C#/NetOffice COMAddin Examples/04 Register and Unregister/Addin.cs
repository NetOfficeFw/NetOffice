using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Tools;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;
using NetOffice.OutlookApi.Tools;
/*
   Register Addin Example
*/
namespace Outlook04AddinCS4
{
    [COMAddin("Outlook04 Sample Addin CS4", "Register Addin Example", LoadBehavior.LoadAtStartup)]
    [ProgId("Outlook04AddinCS4.Connect"), Guid("AFEB261F-A4CD-4060-A11D-C8C456102BD5"), Codebase, Timestamp]
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
            Office.Tools.Contribution.DialogUtils.ShowRegisterError("Outlook04AddinCS4", methodKind, exception);
        }
    }
}