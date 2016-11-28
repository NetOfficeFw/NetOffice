using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;

namespace NetOffice.Tools
{
    /// <summary>
    /// Specify possible registry locations
    /// </summary>
    public enum RegistrySaveLocation
    {        
        /// <summary>
        /// Based on current scope but related addin key want set always in CurrentUser
        /// </summary>
        InstallScopeCurrentUser = 0,

        /// <summary>
        /// Based on current scope
        /// </summary>
        InstallScope = 1,


        /// <summary>
        /// CurrentUser Key
        /// </summary>
        CurrentUser = 2,

        /// <summary>
        /// LocalMachineKey (permissions required)
        /// </summary>
        LocalMachine = 3,
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
        /// Returns info the current combination of RegistryLocation and InstallScope means system key for component register
        /// </summary>
        /// <param name="scope">scope target</param>
        /// <returns>true if machine otherwise false</returns>
        public bool IsMachineComponentTarget(InstallScope scope)
        {
            if (Value == RegistrySaveLocation.InstallScope)
                return scope == InstallScope.System;
            else
                return Value == RegistrySaveLocation.LocalMachine;
        }

                /// <summary>
        /// Returns info the current combination of RegistryLocation and InstallScope means system key for component register
        /// </summary>
        /// <param name="scope">scope target</param>
        /// <returns>true if machine otherwise false</returns>
        public bool IsMachineAddinTarget(InstallScope scope)
        {
            if (Value == RegistrySaveLocation.InstallScopeCurrentUser)
                return false;

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
        /// <param name="createTimeStamp">create timestamp</param>
        public static void CreateApplicationKey(bool isSystem, string officeKey, string progId, int loadBehavior, string friendlyName, string description, bool createTimeStamp)
        {
            CreateApplicationKey(isSystem, officeKey, progId, loadBehavior, friendlyName, description, -1, createTimeStamp);
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
        /// <param name="commandLineSafe">addin is safe for commandline</param>
        /// <param name="createTimeStamp">create timestamp</param>
        public static void CreateApplicationKey(bool isSystem, string officeKey, string progId, int loadBehavior, string friendlyName, string description, int commandLineSafe, bool createTimeStamp)
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
            else
                applicationKey.DeleteValue("CommandLineSafe", false);

            if (createTimeStamp)
                applicationKey.SetValue("CreatedAt", DateTime.Now.ToString(), RegistryValueKind.String);

            applicationKey.Close();           
        }

        /// <summary>
        /// Deletes an office addin key entry
        /// </summary>
        /// <param name="isSystem">install to the system or current user</param>
        /// <param name="officeKey">the office application root key without hive key</param>
        /// <param name="progId">addin progid</param>
        /// <returns>true if no exception occurs</returns>
        public static bool TryDeleteApplicationKey(bool isSystem, string officeKey, string progId)
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
            catch
            {
                return false;
            }
        }
    }
}
