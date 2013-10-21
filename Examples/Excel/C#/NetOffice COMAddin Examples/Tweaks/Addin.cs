using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text;
using System.Runtime.InteropServices;
using NetOffice.Tools;
using NetOffice.ExcelApi.Tools;
using NetOffice.ExcelApi;

/*
 *    This project shows you the Tweak attribute in NetOffice.
 *    You can set the Tweak attribute to set/manipulate NetOffice options or your own options at runtime.
 *    This can be very helpful for developers, may troubleshooting or diagnostics, whatever.
 *    All Tweak settings has to be stored as string value in the current office addin registry key. For example:(HKEY_CurrentUser\Sofware\Microsoft\Office\%Application%\Addins\YourAddin)
 *    You find all possible NetOffice default tweak settings here: http://netoffice.codeplex.com/wikipage?=Tweaks"
 *    In this project you learn how you get control about tweaks and implement your own tweaks in an easy way.*    
 */

namespace NetOfficeTools.TweaksCS4
{
    [COMAddin("NetOfficeCS4 Sample Excel Addin", "This Addin shows you the COMAddin tweak option from the NetOffice Tools", 3)]
    [Guid("DF2DA04E-CD24-4F48-B7F2-A7C3C56E877A"), ProgId("TweakExcelCS4.Addin"), Tweak(true)]  // <== the tweak attribute
    public class Addin : COMAddin
    {
        public Addin()
        {

        }

        // This method was called for all (currently found) tweaks while startup. This means the NetOffice tweaks and your own tweaks.
        // You have to decide the tweak is allowed or not. Please keep in your mind: All NetOffice tweak names starts with 'NO'
        protected override bool AllowApplyTweak(string name, string value)
        {
            // we accept all tweaks
            return true;
        }

        // This method was called from IExtensibility2.OnStartupComplete for all your custom tweaks if its allowed(see AllowApplyTweak)
        protected override void ApplyCustomTweak(string name, string value)
        {
            if (name == "ShowMessageBoxAtStartUp" && value == "yes")
                MessageBox.Show("The tweak sample addin has been loaded.", "Custom Tweak", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // This method was called from IExtensibility2.OnDisconnection for all your allowed custom aplied tweaks to remove or unload them.
        // Please keep in your mind: the method is never called in state of unexpected termination. you have no warranties for the method.
        protected override void DisposeCustomTweak(string name, string value)
        {

        }

        // We set some default- and custom tweaks in the register method.
        // Please note: Installers like .msi or other doesnt call the static register methods for your (managed) addin while un-/registration.
        // You have to set these entries at hand in the corresponding deployment project.
        [RegisterFunction(RegisterMode.CallAfter)]
        public static void Register(Type type, RegisterCall registerCall)
        {
            // SetTweakPersistenceEntry sets the key for you in the current registry key.
            // We set a Netoffice default tweak and a custom tweak.
            SetTweakPersistenceEntry(type, "NOConsoleMode", "trace", false);
            SetTweakPersistenceEntry(type, "ShowMessageBoxAtStartUp", "yes", false);
        }
    }
}
