using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;

namespace NetOffice.DeveloperToolbox.Utils.Registry
{
    public enum UtilsRegistryEntryType
    {
        Normal = 0,
        Faked = 1,
        Default = 2
    }

    public class UtilsRegistryEntries : IEnumerable<UtilsRegistryEntry>
    {
        #region Fields

        private UtilsRegistryKey _parent;

        #endregion

        #region Construction

        protected UtilsRegistryEntries()
        {

        }

        internal UtilsRegistryEntries(UtilsRegistryKey parent)
        {
            _parent = parent;
        }

        #endregion

        #region Properties

        public int Count
        {
            get
            {
                return _parent.InnerKey.ValueCount;
            }
        }
        public UtilsRegistryEntry FakedDefaultKey
        {
            get
            {
                UtilsRegistryEntry entry = new UtilsRegistryEntry(_parent, "", UtilsRegistryEntryType.Faked);
                return entry;
            }
        }

        public UtilsRegistryEntry this[string name]
        {
            get
            {
                UtilsRegistryEntry entry = null;
                if (name == "")
                    entry = new UtilsRegistryEntry(_parent, name, UtilsRegistryEntryType.Default);
                else
                    entry = new UtilsRegistryEntry(_parent, name);

                return entry;
            }
        }

        #endregion

        #region IEnumerable<UtilsRegistryEntry>

        public IEnumerator<UtilsRegistryEntry> GetEnumerator()
        {
            RegistryKey key = _parent.Open();
            string[] names = key.GetValueNames();
            names = SortArray(names);
            foreach (string item in names)
                yield return this[item];
            key.Close();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            RegistryKey key = _parent.Open();
            string[] names = key.GetValueNames();
            names = SortArray(names);
            foreach (string item in names)
                yield return this[item];
            key.Close();
        }

        #endregion

        #region Helper

        private static string ByteArrayToString(byte[] arr)
        {
            System.Text.UnicodeEncoding enc = new System.Text.UnicodeEncoding();
            return enc.GetString(arr);
        }

        private static string[] SortArray(string[] array)
        {
            List<string> list = new List<string>();
            foreach (string item in array)
            {
                if (string.IsNullOrEmpty(item))
                {
                    //found = true;
                    list.Insert(0, item);
                }
                else
                    list.Add(item);
            }
            return list.ToArray();
        }

        #endregion
   }
}
