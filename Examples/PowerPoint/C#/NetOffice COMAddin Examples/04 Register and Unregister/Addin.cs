using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Tools;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using PowerPoint = NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Enums;
using NetOffice.PowerPointApi.Tools;
/*
   Register Addin Example
*/
namespace PowerPoint04AddinCS4
{
    [COMAddin("PowerPoint04 Sample Addin CS4", "Register Addin Example", LoadBehavior.LoadAtStartup)]
    [ProgId("PowerPoint04AddinCS4.Connect"), Guid("A5FD2DA0-8F65-4AB1-862A-B9D01D65ECCB"), Codebase, Timestamp]
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
            Office.Tools.Contribution.DialogUtils.ShowRegisterError("PowerPoint04AddinCS4", methodKind, exception);
        }
    }
}