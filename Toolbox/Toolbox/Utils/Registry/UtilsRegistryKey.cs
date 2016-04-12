using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;

namespace NetOffice.DeveloperToolbox.Utils.Registry
{
    public class UtilsRegistryKey
    {
        #region Fields

        private string          _path;
        private UtilsRegistry   _root;
        private RegistryKey     _innerKey;

        #endregion

        #region Construction

        internal UtilsRegistryKey(UtilsRegistry root, string fullPath)
        {
            _root = root;
            _path = fullPath;
            _innerKey = _root.HiveKey.OpenSubKey(fullPath);
            _innerKey.Close();
        }

        internal UtilsRegistryKey(UtilsRegistry root, RegistryKey innerKey, string path)
        {
            _root = root;
            _innerKey = innerKey;
            _path = path;
        }

        private UtilsRegistryKey Parent
        {
            get
            {
                if (_root.Key.Path == this.Path)
                    return null;

                int position = _path.LastIndexOf("\\", StringComparison.InvariantCultureIgnoreCase);
                if (-1 == position)
                {
                    string parentPath = _root.Key.InnerKey.ToString();
                    RegistryKey key = _root.Key.InnerKey.OpenSubKey(_path);
                    UtilsRegistryKey parentKey = new UtilsRegistryKey(_root, key, parentPath);
                    key.Close();
                    return parentKey;
                }
                else
                {
                    string parentPath = _path.Substring(0, position);
                    RegistryKey key = _root.HiveKey.OpenSubKey(parentPath);
                    UtilsRegistryKey parentKey = new UtilsRegistryKey(_root, key, parentPath);
                    key.Close();
                    return parentKey;
                }
            }
        }

        #endregion

        #region Properties

        internal RegistryKey InnerKey
        {
            get
            {
                return _innerKey;
            }
        }

        internal UtilsRegistry Root
        {
            get
            {
                return _root;
            }
        }

        public string Path
        {
            get
            {
                return _path;
            }
        }

        public string Name
        {
            get
            {
                int postion = _path.LastIndexOf("\\", StringComparison.InvariantCultureIgnoreCase);
                return _path.Substring(postion + 1);
            }
            set
            {
                if ((value != Name) && (value != null))
                { 
                    RegistryKey parentKey = Parent.Open(true);
                    RegistryKey key = Open(true);
                    _innerKey = UtilsRegistry.RenameSubKey(parentKey, Name, value);
                    _path = UtilsRegistry.ReCalculatePath(_path, value);
                    parentKey.Close();
                }
            }
        }

        public UtilsRegistryKeys Keys
        {
            get
            {
                return new UtilsRegistryKeys(this);
            }
        }

        public UtilsRegistryEntries Entries
        {
            get
            {
                return new UtilsRegistryEntries(this);
            }
        }
       
        #endregion

        #region Methods
        
        private string GetNewSubKeyName()
        {
            string name = "#Key";
            UtilsRegistryKeys keys = Keys;
            string[] existingNames = new string[keys.Count];
            int i = 0;
            foreach (UtilsRegistryKey item in keys)
            {
                existingNames[i] = item.Name;
                i++;
            }
            i = 0;

            foreach (string item in existingNames)
            {
                i++;
                if (name.Equals(item, StringComparison.InvariantCultureIgnoreCase))
                    name += i;                
            }
         
            return name;
        }

        public void CreateNewSubKey()
        {
            string name = GetNewSubKeyName();
            string path = _path;
            if (path.StartsWith("HKEY_CURRENT_USER"))
            {
                path = path.Substring("HKEY_CURRENT_USER\\".Length);
                RegistryKey key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(path + "\\" + name);
                key.Close();
            }
            else
            {
                RegistryKey regKey = Open(true);
                RegistryKey key = regKey.CreateSubKey(name);
                key.Close();
                regKey.Close();

            }
        }

        public void Delete()
        {
            _root.HiveKey.DeleteSubKeyTree(Path);
        }

        public RegistryKey Open(bool writable)
        {
            RegistryKey key = _root.HiveKey.OpenSubKey(Path, writable);
            return key;
        }

        public RegistryKey Open()
        {
            RegistryKey key = _root.HiveKey.OpenSubKey(Path);
            return key;
        }

        #endregion

        #region Overrides

        public override string ToString()
        {
            return String.Format("UtilsRegistryKey {0}", Name);
        }

        #endregion
    }
}
