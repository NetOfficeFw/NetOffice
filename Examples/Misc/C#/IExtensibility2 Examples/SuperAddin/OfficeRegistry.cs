using System;
using Microsoft.Win32;
using System.Collections.Generic;
using System.Text;

namespace Super2AddinCS4
{
    /// <summary>
    /// office addin keys and create/delete methods
    /// </summary>
    public static class OfficeRegistry
    {
        // office addin registry keys 

        public static readonly string Office     = "Software\\Microsoft\\Office\\";
        public static readonly string Excel      = "Software\\Microsoft\\Office\\Excel\\AddIns\\";
        public static readonly string Word       = "Software\\Microsoft\\Office\\Word\\AddIns\\";
        public static readonly string Outlook    = "Software\\Microsoft\\Office\\Outlook\\AddIns\\";
        public static readonly string PowerPoint = "Software\\Microsoft\\Office\\PowerPoint\\AddIns\\";
        public static readonly string Access     = "Software\\Microsoft\\Office\\Access\\AddIns\\";

        /// <summary>
        /// creates addin key
        /// </summary>
        /// <param name="key"></param>
        public static void CreateAddinKey(string officeApp, string progId, string name, string description)
        {
            string regKey = GetRegistryKey(officeApp);

            RegistryKey rk = Registry.CurrentUser.CreateSubKey(regKey + progId);
            rk.Close();
            rk = Registry.CurrentUser.OpenSubKey(regKey + progId, true);
            rk.SetValue("LoadBehavior", Convert.ToInt32(3));
            rk.SetValue("FriendlyName", name);
            rk.SetValue("Description", description);
        }

        /// <summary>
        /// deletes addin key
        /// </summary>
        /// <param name="key"></param>
        public static void DeleteAddinKey(string key)
        {
            Registry.CurrentUser.DeleteSubKey(key, false);
        }

        /// <summary>
        /// saves an error in the registry. Messageboxes in ITDExtensibility, IRibbonExtensibility or ICustomtaskPaneConsumer methods cause trouble
        /// </summary>
        /// <param name="officeApp"></param>
        /// <param name="_progId"></param>
        /// <param name="message"></param>
        /// <param name="exception"></param>
        public static void LogErrorMessage(string officeApp, string progId, string message, Exception exception)
        {
            string regKey = GetRegistryKey(officeApp);

            RegistryKey rk = Registry.CurrentUser.OpenSubKey(regKey + progId, true);

            rk.SetValue("ErrorTimestamp", DateTime.Now.ToString());
            rk.SetValue("ErrorMessage", message);
            rk.SetValue("ErrorException", exception.Message);
            rk.Close();
        }

        private static string GetRegistryKey(string officeApp)
        {
            switch (officeApp)
            { 
                case "Excel":
                    return Excel;
                case "Word":
                    return Word;
                case "Outlook":
                    return Outlook;
                case "PowerPoint":
                    return PowerPoint;
                case "Access":
                    return Access;
                default :
                    throw new ArgumentOutOfRangeException("officeApp");
            }
        }
    }
}
