using System;
using Microsoft.Win32;
using System.Collections.Generic;
using System.Text;

namespace SuperAddin
{
    /// <summary>
    /// office addin keys and create/delete methods
    /// </summary>
    public static class OfficeRegistry
    {
        /*
           office addin registry keys 
        */

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
        public static void CreateAddinKey(string key)
        {
            RegistryKey regKey = Registry.CurrentUser.CreateSubKey(key);
            regKey.Close();
            regKey = Registry.CurrentUser.OpenSubKey(key, true);
            regKey.SetValue("LoadBehavior", Convert.ToInt32(3));
            regKey.SetValue("FriendlyName", "SuperAddin");
            regKey.SetValue("Description", "example for versionindependent addin loaded in all office products");
        }


        /// <summary>
        /// deletes addin key
        /// </summary>
        /// <param name="key"></param>
        public static void DeleteAddinKey(string key)
        {
            Registry.CurrentUser.DeleteSubKey(key);
        }
    }
}
