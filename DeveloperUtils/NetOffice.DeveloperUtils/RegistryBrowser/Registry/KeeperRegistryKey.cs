using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;

namespace NetOffice.DeveloperUtils.RegistryBrowser
{
    public class KeeperRegistryKey
    {
        #region Member

        private string          _key;
        private string          _name;

        private KeeperRegistryEntries _entries = null;
        private KeeperRegistryKeys    _subKeys = null;

        #endregion

        #region Properties

        public string Name
        {
            get
            {
                return _name;
            }
        }

       
        public KeeperRegistryEntries Entries
        {
            get
            {
                return _entries;
            }
        }
       

        public KeeperRegistryKeys Keys
        {
            get
            {
                return _subKeys;
            }
        }
       
        #endregion

        #region Construction

        internal KeeperRegistryKey(RegistryKey root, string rootKey)
        {
            _key = rootKey;
            _name = rootKey.Substring(rootKey.LastIndexOf(@"\")+1);

            _entries = new KeeperRegistryEntries(root, _key);
            _subKeys = new KeeperRegistryKeys(root, _key);
        }

        #endregion
    }
}
