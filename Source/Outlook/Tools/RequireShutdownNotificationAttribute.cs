using System;
using Microsoft.Win32;

namespace NetOffice.OutlookApi.Tools
{
    /// <summary>
    /// Signalize to Outlook that the Addin want have Shutdown events
    /// See https://msdn.microsoft.com/library/office/ee720183.aspx
    /// </summary>
    public class RequireShutdownNotificationAttribute : System.Attribute
    {
        /// <summary>
        /// Try get attribute from type
        /// </summary>
        /// <param name="type">the type you want looking for the attribute</param>
        /// <returns>RequireShutdownNotificationAttribute or null</returns>
        public static RequireShutdownNotificationAttribute GetAttribute(Type type)
        {
            object[] array = type.GetCustomAttributes(typeof(RequireShutdownNotificationAttribute), false);
            if (array.Length == 0)
                return null;
            else
                return array[0] as RequireShutdownNotificationAttribute;
        }

        /// <summary>
        /// Creates the RequireShutdownNotification registry value
        /// </summary>
        /// <param name="isSystem">install to the system or current user</param>
        /// <param name="officeKey">the office application root key without hive key</param>
        /// <param name="progId">addin progid</param>
        public static void CreateApplicationKey(bool isSystem, string officeKey, string progId)
        {
            string targetKey = officeKey + progId;
            RegistryKey applicationKey = null;
            if (isSystem)
                applicationKey = Registry.LocalMachine.CreateSubKey(targetKey);
            else
                applicationKey = Registry.CurrentUser.CreateSubKey(targetKey);

            applicationKey.Close();

            if (isSystem)
                applicationKey = Registry.LocalMachine.OpenSubKey(targetKey, true);
            else
                applicationKey = Registry.CurrentUser.OpenSubKey(targetKey, true);

            applicationKey.SetValue("RequireShutdownNotification", 1, RegistryValueKind.DWord);
            applicationKey.Close();
        }
    }
}
