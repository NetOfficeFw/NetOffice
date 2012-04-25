using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;

namespace NetOffice.DeveloperToolbox.WindowsRegistry
{
    public class UtilsRegistry
    {
        #region Constants

        private RegistryKey _hiveKey;
        private RegistryKey _innerKey;
        private string      _path;

        #endregion

        #region Member

        private UtilsRegistryKey     _key;
        private UtilsRegistryEntries _entries;

        #endregion

        #region Construction

        public UtilsRegistry(RegistryKey hiveKey, string path)
        {
            _hiveKey = hiveKey;
            _path = path;
            _innerKey = hiveKey.OpenSubKey(path);
        }

        #endregion
     
        #region Properties

        public bool Exists
        {
            get
            {
                bool retValue = false;
                RegistryKey rk = _hiveKey.OpenSubKey(_path);
                if (rk != null)
                {
                    rk.Close();
                    retValue = true;
                }
                return retValue;
            }
        }
        
        public RegistryKey HiveKey
        {
            get
            {
                return _hiveKey;
            }
        }

        public RegistryKey InnerKey
        {
            get
            {
                return _hiveKey;
            }
        }
         
        public UtilsRegistryKey Key
        {
            get
            {
                if (null == _key)
                    _key = new UtilsRegistryKey(this, _innerKey, _path);
                return _key;
            }
        }

        public UtilsRegistryEntries Entries
        {
            get
            {
                if (null == _entries)
                    _entries = new UtilsRegistryEntries(Key);
                return _entries;
            }
        }

        #endregion

        #region Internal Statics

        internal static string ReCalculatePath(string path, string newKeyName)
        {
            int position = path.LastIndexOf("\\", StringComparison.InvariantCultureIgnoreCase);
            string cutKey = path.Substring(0, position);
            return cutKey + newKeyName;
        }

        internal static RegistryKey RenameSubKey(RegistryKey parentKey, string subKeyName, string newSubKeyName)
        {
            RegistryKey newKey = CopyKey(parentKey, subKeyName, newSubKeyName);
            parentKey.DeleteSubKeyTree(subKeyName);
            return newKey;
        }

        public static RegistryKey CopyKey(RegistryKey parentKey, string keyNameToCopy, string newKeyName)
        {
            RegistryKey destinationKey = parentKey.CreateSubKey(newKeyName);
            RegistryKey sourceKey = parentKey.OpenSubKey(keyNameToCopy);
            RecurseCopyKey(sourceKey, destinationKey);
            return destinationKey;
        }

        private static void RecurseCopyKey(RegistryKey sourceKey, RegistryKey destinationKey)
        {
            foreach (string valueName in sourceKey.GetValueNames())
            {
                object objValue = sourceKey.GetValue(valueName);
                RegistryValueKind valKind = sourceKey.GetValueKind(valueName);
                destinationKey.SetValue(valueName, objValue, valKind);
            }

            foreach (string sourceSubKeyName in sourceKey.GetSubKeyNames())
            {
                RegistryKey sourceSubKey = sourceKey.OpenSubKey(sourceSubKeyName);
                RegistryKey destSubKey = destinationKey.CreateSubKey(sourceSubKeyName);
                RecurseCopyKey(sourceSubKey, destSubKey);
            }
        }

        #endregion
    }
}
