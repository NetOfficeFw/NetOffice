using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;

namespace NetOffice.Tools
{
    /// <summary>
    /// COMAddin Register method want create Lockback Bypass Key - see: http://support.microsoft.com/kb/948461
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Class, AllowMultiple = false)]
    public class LockbackAttribute : System.Attribute
    {
        /// <summary>
        /// Creates Office .NET Framework Lockback Bypass Key
        /// </summary>
        /// <param name="isSystem">ClassesRoot want used or CurrentUser</param>
        /// <returns>true if created otherwise false</returns>
        public static bool CreateKey(bool isSystem)
        {
            try
            {
                RegistryKey lockbackKey = null;
                if(isSystem)
                    lockbackKey = Registry.ClassesRoot.CreateSubKey("Interface\\{000C0601-0000-0000-C000-000000000046}");
                else
                    lockbackKey = Registry.CurrentUser.CreateSubKey("Software\\Classes\\Interface\\{000C0601-0000-0000-C000-000000000046}");

                string defaultValue = lockbackKey.GetValue("") as string;
                if (null == defaultValue)
                    lockbackKey.SetValue("", "Office .NET Framework Lockback Bypass Key");
                lockbackKey.Close();
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
