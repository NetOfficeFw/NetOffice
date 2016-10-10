using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;

namespace NetOffice.Tools
{
    /// <summary>
    /// specifiy possible registry locations
    /// </summary>
    public enum RegistrySaveLocation
    {       
        /// <summary>
        /// Based on current scope
        /// </summary>
        InstallScope = 0,

        /// <summary>
        /// CurrentUser Key
        /// </summary>
        CurrentUser = 1,

        /// <summary>
        /// LocalMachineKey (permissions required)
        /// </summary>
        LocalMachine = 2,
    }

    /// <summary>
    /// Specify the addin registry keys for office was created in the Machine key or current user
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Class, AllowMultiple = false)]
    public class RegistryLocationAttribute : System.Attribute
    {
        /// <summary>
        /// Registry Location
        /// </summary>
        public readonly RegistrySaveLocation Value;

        /// <summary>
        /// Creates an instance of the attribute
        /// </summary>
        /// <param name="value">Registry location</param>
        public RegistryLocationAttribute(RegistrySaveLocation value)
        {
            Value = value;
        }

        /// <summary>
        /// Returns info the current combination of RegistryLocation and InstallScope means system key
        /// </summary>
        /// <param name="scope">scope target</param>
        /// <returns>true if machine otherwise false</returns>
        public bool IsMachineTarget(InstallScope scope)
        {
            if (Value == RegistrySaveLocation.InstallScope)
                return scope == InstallScope.System;
            else
                return Value == RegistrySaveLocation.LocalMachine;
        }

        /// <summary>
        /// Creates the office application registry to load the addin
        /// </summary>
        /// <param name="isSystem">install to the system or current user</param>
        /// <param name="officeKey">the office application root key without hive key</param>
        /// <param name="progId">addin progid</param>
        /// <param name="loadBehavior">addin load behaviour</param>
        /// <param name="friendlyName">addin caption</param>
        /// <param name="description">addin detailed description</param>
        public static void CreateApplicationKey(bool isSystem, string officeKey, string progId, int loadBehavior, string friendlyName, string description)
        {
            CreateApplicationKey(isSystem, officeKey, progId, loadBehavior, friendlyName, description, -1);
        }

        /// <summary>
        /// Creates the office application registry to load the addin
        /// </summary>
        /// <param name="isSystem">install to the system or current user</param>
        /// <param name="officeKey">the office application root key without hive key</param>
        /// <param name="progId">addin progid</param>
        /// <param name="loadBehavior">addin load behaviour</param>
        /// <param name="friendlyName">addin caption</param>
        /// <param name="description">addin detailed description</param>
        /// <param name="commandLineSafe">addin is commandline safe</param>
        public static void CreateApplicationKey(bool isSystem, string officeKey, string progId, int loadBehavior, string friendlyName, string description, int commandLineSafe)
        {
            string targetKey = officeKey + progId;
            RegistryKey applicationKey = null;
            if(isSystem)
                applicationKey = Registry.LocalMachine.CreateSubKey(targetKey);
            else
                applicationKey = Registry.CurrentUser.CreateSubKey(targetKey);

            applicationKey.Close();

            if (isSystem)
                applicationKey = Registry.LocalMachine.OpenSubKey(targetKey, true);
            else
                applicationKey = Registry.CurrentUser.OpenSubKey(targetKey, true);

            applicationKey.SetValue("LoadBehavior", loadBehavior, RegistryValueKind.DWord);
            applicationKey.SetValue("FriendlyName", friendlyName, RegistryValueKind.String);
            applicationKey.SetValue("Description", description, RegistryValueKind.String);
            if (-1 != commandLineSafe)
                applicationKey.SetValue("CommandLineSafe", commandLineSafe, RegistryValueKind.DWord);

            applicationKey.Close();
        }

        /// <summary>
        /// Deletes an office addin key entry
        /// </summary>
        /// <param name="isSystem">install to the system or current user</param>
        /// <param name="officeKey">the office application root key without hive key</param>
        /// <param name="progId">addin progid</param>
        /// <returns>true if no exception occurs</returns>
        public static bool DeleteApplicationKey(bool isSystem, string officeKey, string progId)
        {
            try
            {             
                if (isSystem)
                {
                    Registry.LocalMachine.DeleteSubKey(officeKey + progId, false);
                }
                else
                {
                    Registry.CurrentUser.DeleteSubKey(officeKey + progId, false);
                }
                return true;
            }
            catch (Exception)
            {

                return false;
            }
        }
    }
}
