using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;

namespace NetOffice.DeveloperUtils.RegistryBrowser
{ 
    public class KeeperRegistryEntries
    {       
        #region Member

        private string _key = "";
        private List<KeeperRegistryEntry> _list;

        #endregion

        #region Construction

        internal KeeperRegistryEntries(RegistryKey root, string key)
        {
            _key = key;
            _list = new List<KeeperRegistryEntry>();

            RegistryKey rk = root.OpenSubKey(_key, false);
            if (null != rk)
            {
                string[] values = rk.GetValueNames();
                foreach (string value in values)
                {
                    KeeperRegistryEntry entry = null;
                    RegistryValueKind rvk = rk.GetValueKind(value);
                    object o = rk.GetValue(value);
                    bool isBinary = false;

                    if (o is byte[])
                    {
                        o = ByteArrayToString((byte[])o);
                        isBinary = true;
                    }
                    else
                        o = rk.GetValue(value);

                    entry = new KeeperRegistryEntry(root, key, value, o, rvk, isBinary);
                    _list.Add(entry);
                }
                rk.Close();
            }
        }

        #endregion

        private static string ByteArrayToString(byte[] arr)
        {
            System.Text.ASCIIEncoding enc = new System.Text.ASCIIEncoding();
            return enc.GetString(arr);
        }

        #region Properties

        public int Count
        {
            get 
            {
                return _list.Count;
            }
        }

        public KeeperRegistryEntry this[int i]
        {
            get 
            {
                return _list[i];
            }
        }

        public KeeperRegistryEntry this[string Name]
        {
            get
            {
                int iCount = Count;
                for (int i = 1; i <= iCount; i++)
                {
                    KeeperRegistryEntry entry = this[i - 1];
                    if (Name.Equals(entry.Name, StringComparison.CurrentCultureIgnoreCase) == true)
                        return entry;
                }

                throw (new IndexOutOfRangeException("RegistryEntry " + Name + " not found."));
            }
        }

        /// <summary>
        /// Foreach Enumerator
        /// </summary>
        /// <returns></returns>
        public IEnumerator GetEnumerator()
        {
            int iCount = Count;
            KeeperRegistryEntry[] res_entries = new KeeperRegistryEntry[iCount];

            for (int i = 0; i < iCount; i++)
                res_entries[i] = this[i];

            for (int i = 0; i < res_entries.Length; i++)
                yield return res_entries[i];
        }
 
        #endregion
    }
}
