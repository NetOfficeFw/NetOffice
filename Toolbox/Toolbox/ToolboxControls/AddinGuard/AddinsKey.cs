using System;
using System.Collections.Generic;
using Microsoft.Win32;
using System.Text;

namespace NetOffice.DeveloperToolbox.ToolboxControls.AddinGuard
{
    class AddinsKey
    {
        #region Fields

        WatchController _parent;
        List<AddinKey> _addinList;
        string _name;

        int _subKeyCount;
        string[] _subKeyNames = new string[0];

        RegistryKey _rootKey;
        string _registryPath;
        
        #endregion
 
        #region Properties
         
        public AddinKey[] Addins
        {
            get
            {
                RegistryKey key = _rootKey.OpenSubKey(_registryPath);
                if (null != key)
                {
                    string[] subKeyNames = key.GetSubKeyNames();
                    foreach (string subKeyName in subKeyNames)
                    {
                        if (!Contains(_addinList, subKeyName))
                        { 
                            AddinKey subItem = new AddinKey(_parent, this,_registryPath + "\\" + subKeyName, subKeyName);
                            _addinList.Add(subItem);
                        }
                    }
                    key.Close();
                    DeleteNotExistingAddins(_addinList, subKeyNames);
                    return _addinList.ToArray();
                }
                return new AddinKey[0];
            }
        }

        public string Name
        {
            get
            {
                return _name;
            }
        }
 
        public int SubKeyCount
        {
            get 
            {
                return _subKeyCount;
            }
        }

       
        public RegistryKey RootKey
        {
            get 
            {
                return _rootKey;
            }
        }

        public string RegistryPath
        {
            get
            {
                return _registryPath;
            }
        }
       
        #endregion

        #region Construction

        public AddinsKey(WatchController parent,string name, RegistryKey rootKey, string registryPath)
        {
            _parent = parent;
            _rootKey = rootKey;
            _name = name;
            _registryPath = registryPath;
            _addinList = new List<AddinKey>();
            GetKeySubKeys();
        }

        #endregion

        #region Methods

        internal NotifyKind CheckChangedSubKeys(ref RegistryChangeInfo changeInfo)
        {
            RegistryKey key = _rootKey.OpenSubKey(_registryPath);
            if (null != key)
            {
                string[] subKeyNames = key.GetSubKeyNames();
                int subKeyCount = key.SubKeyCount;
                if(subKeyCount != _subKeyCount)
                {
                    int oldSubKeyCount = _subKeyCount;
                    _subKeyCount = subKeyCount;
                    key.Close();
                    _parent.RaisePropertyChanged(this);

                    NotifyKind returnKind = NotifyKind.Nothing;
                    if ((subKeyCount > oldSubKeyCount) && (!_parent.FirstRun))
                    {
                        AddinSubkeysIncrementInfo newKeyInfo = new AddinSubkeysIncrementInfo();
                        newKeyInfo.RootKey = _rootKey;
                        newKeyInfo.KeyPath = _registryPath;
                        newKeyInfo.KeyName =  GetNewValueName(subKeyNames, _subKeyNames);
                        changeInfo = newKeyInfo;
                        _subKeyNames = subKeyNames;
                        returnKind = NotifyKind.AddinSubKeysIncrement;
                    }
                    else if(!_parent.FirstRun)
                    {
                        AddinSubkeysDecrementInfo deleteKeyInfo = new AddinSubkeysDecrementInfo();
                        deleteKeyInfo.RootKey = _rootKey;
                        deleteKeyInfo.KeyPath = _registryPath;
                        deleteKeyInfo.KeyName = GetDeletedValueName(subKeyNames, _subKeyNames);
                        changeInfo = deleteKeyInfo;
                        _subKeyNames = subKeyNames;
                        returnKind = NotifyKind.AddinSubKeysDecrement;
                    }
                    return returnKind;
                }
                else
                {
                    NotifyKind returnKind = NotifyKind.Nothing;
                    foreach (string item in subKeyNames)
                    {
                        if (!Contains(_subKeyNames, item))
                        {
                            _parent.RaisePropertyChanged(this);
                            if (!_parent.FirstRun)
                            { 
                                AddinSubkeyNameChangedInfo nameInfo = new AddinSubkeyNameChangedInfo();
                                nameInfo.RootKey = _rootKey;
                                nameInfo.KeyPath = _registryPath;
                                nameInfo.OldKeyName = item;
                                nameInfo.NewKeyName = GetChangedValueName(subKeyNames, _subKeyNames);
                                changeInfo = nameInfo;
                                _subKeyNames = subKeyNames;
                                returnKind = NotifyKind.AddinSubKeyNameChanged;
                            }
                            break;
                        }
                    }
                    _subKeyNames = subKeyNames;
                    key.Close();
                    return returnKind;
                }
            }
            return NotifyKind.Nothing;
        }
 
        #endregion

        #region Static Methods

        private static string GetChangedValueName(string[] newValueNames, string[] oldValueNames)
        {
            string item = "";
            foreach (string newValue in newValueNames)
            {
                bool found = false;
                foreach (string oldValue in oldValueNames)
                {
                    if (newValue == oldValue)
                    {
                        found = true;
                        break;
                    }
                    else
                        item = oldValue;
                }
                if (!found)
                    return item;
            }
            throw new ArgumentException("No changed Name found");
        }

        private static string GetDeletedValueName(string[] newValueNames, string[] oldValueNames)
        {
            foreach (string oldValue in oldValueNames)
            {
                bool found = false;
                foreach (string newValue in newValueNames)
                {
                    if (newValue == oldValue)
                    {
                        found = true;
                        break;
                    }
                }
                if (!found)
                    return oldValue;
            }

            throw new ArgumentException("No deleted Value found");
        }

        private static string GetNewValueName(string[] newValueNames, string[] oldValueNames)
        {
            foreach (string newValue in newValueNames)
            {
                bool found = false;
                foreach (string oldValue in oldValueNames)
                {
                    if (newValue == oldValue)
                    {
                        found = true;
                        break;
                    }
                }
                if (!found)
                    return newValue;
            }
            throw new ArgumentException("No new Value found");
        }

        public static RegistryValueKind ConvertStringToRegistryValueKind(string expression)
        {
            string value = expression.Trim();
            switch (value)
            {
                case "String":
                    return RegistryValueKind.String;
                case "ExpandString":
                    return RegistryValueKind.ExpandString;
                case "Binary":
                    return RegistryValueKind.Binary;
                case "DWord":
                    return RegistryValueKind.DWord;
                case "MultiString":
                    return RegistryValueKind.MultiString;
                case "QWord":
                    return RegistryValueKind.QWord;
                default:
                    return RegistryValueKind.Unknown;
            }
        }

        internal static bool Contains(string[] array, string item)
        {
            foreach (string arrayItem in array)
            {
                if (arrayItem == item)
                    return true;
            }
            return false;
        }

        internal static bool IsEqual(object value1, object value2)
        {
            if ((null == value1) && (null == value2))
                return true;

            if (null == value1)
                return false;

            if (null == value2)
                return false;

            string string1 = value1.ToString();
            string string2 = value2.ToString();

            return (string1.Equals(string2, StringComparison.InvariantCultureIgnoreCase));
        }

        #endregion

        #region Private Methods

        private static bool Contains(List<AddinKey> list, string name)
        {
            foreach (AddinKey addinKey in list)
            {
                if (name == addinKey.Name)
                    return true;
            }
            return false;
        }

        private static void DeleteNotExistingAddins(List<AddinKey> list, string[] subKeyNames)
        {
            List<AddinKey> deleteList = new List<AddinKey>();
            foreach (AddinKey itemAddin in list)
            {
                bool found = false;
                foreach (string name in subKeyNames)
                {
                    if (itemAddin.Name == name)
                    { 
                        found = true;
                        break;
                    }
                }
                if (!found)
                    deleteList.Add(itemAddin);
            }

            foreach (AddinKey itemAddin in deleteList)
                list.Remove(itemAddin);
        }

        private void GetKeySubKeys()
        {
            RegistryKey key = _rootKey.OpenSubKey(_registryPath);
            if (null != key)
            {
                _subKeyCount = key.SubKeyCount;
                _subKeyNames = key.GetValueNames();
                key.Close();
            }
        }
  
        #endregion
    }
}
