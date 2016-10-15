using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;

namespace RegAddin.Common
{
    internal class Registry
    {
        private static string _userBypass = "Software\\Classes\\";

        internal void DeleteComponentKey(SingletonSettings.RegisterMode mode, string key)
        {
            RegistryKey rootKey;
            if (mode == SingletonSettings.RegisterMode.System)
                rootKey = Microsoft.Win32.Registry.ClassesRoot;
            else
                rootKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(_userBypass, true);
            
            rootKey.DeleteSubKeyTree(key, false);
            rootKey.Close();
        }

        internal void DeleteComponentKey(SingletonSettings.UnRegisterMode mode, string key)
        {
            RegistryKey[] rootKeys = null;

            switch (mode)
            {
                case SingletonSettings.UnRegisterMode.Auto:
                    rootKeys = new RegistryKey[2];
                    rootKeys[0] = Microsoft.Win32.Registry.ClassesRoot;
                    rootKeys[1] = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(_userBypass, true);
                    break;
                case SingletonSettings.UnRegisterMode.System:
                    rootKeys = new RegistryKey[] { Microsoft.Win32.Registry.ClassesRoot };
                    break;
                case SingletonSettings.UnRegisterMode.User:
                    rootKeys = new RegistryKey[] { Microsoft.Win32.Registry.CurrentUser.OpenSubKey(_userBypass, true) };
                    break;
                default:
                    throw new IndexOutOfRangeException("mode");
            }

            foreach (RegistryKey item in rootKeys)
            {
                try
                {                  
                    item.DeleteSubKeyTree(key, false);
                    item.Close();
                }
                catch (System.Security.SecurityException)
                {
                    if (mode != SingletonSettings.UnRegisterMode.Auto)
                        throw;
                }
                catch (System.UnauthorizedAccessException)
                {
                    if (mode != SingletonSettings.UnRegisterMode.Auto)
                        throw;
                }
                catch (Exception)
                {
                    throw;
                }
            }        
        }

        internal void CreateComponentValue(RegistryKey componentKey, string name, object value, RegistryValueKind kind)
        {
            if (null == componentKey)
                throw new ArgumentNullException("Unable to find specified registry key.");

            if (componentKey.GetValue(name, null) != null && componentKey.GetValueKind(name) != kind)
                componentKey.DeleteValue(name, false);

            componentKey.SetValue(name, value, kind);
        }

        internal void CreateComponentValue(SingletonSettings.RegisterMode mode, string key, string name, object value, RegistryValueKind kind)
        {
            RegistryKey componentKey;
            if (mode == SingletonSettings.RegisterMode.System)
                componentKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(key, true);
            else
                componentKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(_userBypass + key, true);

            if (null == componentKey)
                throw new ArgumentOutOfRangeException("Unable to find specified registry key.");

            CreateComponentValue(componentKey, name, value, kind);
        }

        internal RegistryKey CreateComponentKey(SingletonSettings.RegisterMode mode, params string[] key)
        {
            RegistryKey rootKey;
            if (mode == SingletonSettings.RegisterMode.System)
                rootKey = Microsoft.Win32.Registry.ClassesRoot;
            else
                rootKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(_userBypass, true);

            bool isFirst = true;
            string keys = String.Empty;
            foreach (string item in key)
            {
                if (!isFirst)
                    keys += "\\";
                if (isFirst)
                    isFirst = false;
                keys += item;
            }

            RegistryKey result = rootKey.CreateSubKey(keys);
            rootKey.Close();
            return result;
        }
    }
}
