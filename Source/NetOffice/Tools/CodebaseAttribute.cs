using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// COMAddin Register/Unregister methods want add/remove "Codebase" registry key.
    /// A missing Codebase attribute means a Codebase(true) attribute by default 
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Class, AllowMultiple = false)]
    public class CodebaseAttribute : System.Attribute
    { 
        /// <summary>
        /// Create/Delete Codebase entry. True by default
        /// </summary>
        public readonly bool Value;

        /// <summary>
        /// Creates an instance of the attribute
        /// </summary>
        public CodebaseAttribute()
        {
            Value = true;
        }

        /// <summary>
        /// Creates an instance of the attribute
        /// </summary>
        /// <param name="value">create entry</param>
        public CodebaseAttribute(bool value)
        {
            Value = value;
        }

        /// <summary>
        /// Remove registry codebase value
        /// </summary>
        /// <param name="guid">addin id</param>
        ///  <param name="isSystem">delete in ClassesRoot otherwise CurrentUser</param>
        /// <param name="assemblyVersion">assembly version</param>
        /// <param name="throwExceptionOnError">throw exception on error</param>
        /// <returns>true if removed otherwise false</returns>
        public static bool DeleteValue(Guid guid, bool isSystem, string assemblyVersion, bool throwExceptionOnError)
        {
            try
            {
                Microsoft.Win32.RegistryKey key = TryGetKey(guid, isSystem, assemblyVersion);
                if (null != key && null != assemblyVersion)
                {
                    key.DeleteValue("Codebase", false);
                    return true;
                }
                else
                { 
                    return true;
                }
            }
            catch
            {
                if (throwExceptionOnError)
                    throw;
                return false;
            }
        }

        /// <summary>
        /// Create registry codebase value
        /// </summary>
        /// <param name="guid">addin id</param>
        ///  <param name="isSystem">delete in ClassesRoot otherwise CurrentUser</param>
        /// <param name="assemblyVersion">assembly version</param>
        /// <param name="codebase">given codebase path</param>        
        public static void CreateValue(Guid guid, bool isSystem, string assemblyVersion, string codebase)
        {
            Microsoft.Win32.RegistryKey key = CreateKey(guid, isSystem, assemblyVersion);
            key.SetValue("CodeBase", codebase);
            key.Close();
        }

        /// <summary>
        /// Try to open codebase key
        /// </summary>
        /// <param name="guid">addin id</param>
        /// <param name="isSystem">create in ClassesRoot otherwise CurrentUser</param>
        /// <param name="assemblyVersion">assembly version</param>
        /// <returns>key or null if not found</returns>
        private static Microsoft.Win32.RegistryKey TryGetKey(Guid guid, bool isSystem, string assemblyVersion)
        {
            Microsoft.Win32.RegistryKey result = null;
            if (isSystem)
            {
                result = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey("CLSID\\{" + guid.ToString().ToUpper() + "}\\InprocServer32\\" + assemblyVersion, true);
            }
            else
            {
                result = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\\Classes\\CLSID\\{" + guid.ToString().ToUpper() + "}\\InprocServer32\\" + assemblyVersion, true);
            }
            return result;
        }

        /// <summary>
        /// Create or open registry codebase key
        /// </summary>
        /// <param name="guid">addin id</param>
        /// <param name="isSystem">create in ClassesRoot otherwise CurrentUser</param>
        /// <param name="assemblyVersion">assembly version</param>
        /// <returns>key</returns>
        private static Microsoft.Win32.RegistryKey CreateKey(Guid guid, bool isSystem, string assemblyVersion)
        {
            Microsoft.Win32.RegistryKey result = null;
            if (isSystem)
            {               
                result = Microsoft.Win32.Registry.ClassesRoot.CreateSubKey("CLSID\\{" + guid.ToString().ToUpper() + "}\\InprocServer32\\" + assemblyVersion);                
            }
            else
            {               
                result = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\Classes\\CLSID\\{" + guid.ToString().ToUpper() + "}\\InprocServer32\\" + assemblyVersion);
            }
            return result;
        } 
    }
}
