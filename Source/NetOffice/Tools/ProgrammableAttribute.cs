using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// COMAddin Register/Unregister methods want add/remove "Programmable" registry key
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Class, AllowMultiple = false)]
    public class ProgrammableAttribute : System.Attribute
    {
        /// <summary>
        /// Remove registry programmable keys
        /// </summary>
        /// <param name="guid">addin id</param>
        /// <param name="isSystem">delete in ClassesRoot otherwise CurrentUser</param>
        /// <param name="throwExceptionOnError">throw exception on error</param>
        /// <returns>true if removed otherwise false</returns>
        public static bool DeleteKeys(Guid guid, bool isSystem, bool throwExceptionOnError)
        {
            try
            {
                if (isSystem)
                {
                    Microsoft.Win32.Registry.ClassesRoot.DeleteSubKey(@"CLSID\{" + guid.ToString().ToUpper() + @"}\Programmable", false);
                }
                else
                {
                    Microsoft.Win32.Registry.CurrentUser.DeleteSubKey(@"Software\Classes\CLSID\{" + guid.ToString().ToUpper() + @"}\Programmable", false);
                }
                return true;
            }
            catch
            {
                if(throwExceptionOnError)
                    throw;
                return false;
            }
        }

        /// <summary>
        /// Create registry programmable keys
        /// </summary>
        /// <param name="guid">addin id</param>
        /// <param name="isSystem">create in ClassesRoot otherwise CurrentUser</param>
        public static void CreateKeys(Guid guid, bool isSystem)
        {
            if (isSystem)
            {
                Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.ClassesRoot.CreateSubKey(@"CLSID\{" + guid.ToString().ToUpper() + @"}\Programmable");
                key.Close();
            }
            else
            {
                Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"Software\Classes\CLSID\{" + guid.ToString().ToUpper() + @"}\Programmable");
                key.Close();
            }
        }
    }
}
