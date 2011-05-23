using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;

namespace NetOffice.DeveloperUtils.RegistryBrowser
{
    public class KeeperRegistryKeys
    {
        #region Member

        private string                  _key;
        private List<KeeperRegistryKey> _list;
        
        #endregion
           
        #region Construction

        internal KeeperRegistryKeys(RegistryKey root, string Key)
        {
            _key    = Key;
            _list   = new List<KeeperRegistryKey>();

            RegistryKey rk = root.OpenSubKey(_key, false);
            if (null != rk)
            { 
                string[] Subkeys = rk.GetSubKeyNames();
                foreach (string subKey in Subkeys)
                {
                    KeeperRegistryKey NewKey = new KeeperRegistryKey(root, _key + "\\" + subKey);
                    _list.Add(NewKey);
                }
                rk.Close();
            }
        }

        #endregion

        #region Properties

        public int Count
        {
            get
            {
                return _list.Count;
            }
        }

        public KeeperRegistryKey this[int i]
        {
            get
            {
                return _list[i];
            }
        }

        public KeeperRegistryKey this[string Name]
        {
            get
            {
                int iCount = Count;
                for (int i = 1; i <= iCount; i++)
                {
                    KeeperRegistryKey entry = this[i - 1];
                    if (Name.Equals(entry.Name, StringComparison.CurrentCultureIgnoreCase) == true)
                        return entry;
                }
                return null;
            }
        }

        /// <summary>
        /// Foreach Enumerator
        /// </summary>
        /// <returns></returns>
        public IEnumerator GetEnumerator()
        {
            int iCount = this.Count;
            KeeperRegistryKey[] res_keys = new KeeperRegistryKey[iCount];

            for (int i = 0; i < iCount; i++)
                res_keys[i] = this[i];

            for (int i = 0; i < res_keys.Length; i++)
            {
                yield return res_keys[i];
            }
        }

        #endregion
    }
}
