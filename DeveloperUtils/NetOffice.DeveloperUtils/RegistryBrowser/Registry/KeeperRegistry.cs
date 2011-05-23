using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;

namespace NetOffice.DeveloperUtils.RegistryBrowser
{
    public class KeeperRegistry
    {
        #region Constants

        private RegistryKey _hiveKey;
        private string      _rootKey;

        #endregion

        #region Construction
        
        public KeeperRegistry(RegistryKey hiveKey, string rootKey)
        {
            _hiveKey = hiveKey;
            _rootKey = rootKey;
   
        }

        #endregion

        #region Member

        private KeeperRegistryKey _key;
        private KeeperRegistryEntries _entries;

        #endregion

        #region Properties

        public bool Exists
        {
            get
            {
                bool retValue = false;
                RegistryKey rk = _hiveKey.OpenSubKey(_rootKey, false);
                if (rk != null)
                {
                    rk.Close();
                    retValue = true;
                }
                return retValue;
            }
        }

        public KeeperRegistryKey Key
        {
            get
            {
                if (null == _key)
                    _key = new KeeperRegistryKey(_hiveKey, _rootKey);
                return _key;
            }
        }

        public KeeperRegistryEntries Entries
        {

            get
            {
                if (null == _entries)
                    _entries = new KeeperRegistryEntries(_hiveKey, _rootKey);
                return _entries;
            }
        }

        #endregion
    }
}
