using System;
using System.Collections.Generic;
using Microsoft.Win32;
using System.Text;

namespace NetOffice.DeveloperToolbox.ToolboxControls.AddinGuard
{
    class DisabledKey
    {
        #region Fields

        WatchController _parent;
        string _name;
        int _valueCount;
        int _lastValueCount;
        string[] _valueNames = new string[0];
        Dictionary<string, string> _convertedNames = new Dictionary<string, string>();
        List<DisabledValue> _valuesList = new List<DisabledValue>();
        List<string> _keyExists = new List<string>(); 
        RegistryKey _rootKey;
        string _registryPath;
       
        #endregion
 
        #region Properties

        private static void DeleteNotExistingItems(List<DisabledValue> list, string[] valueNames)
        {
            List<DisabledValue> deleteList = new List<DisabledValue>();
            foreach (DisabledValue itemAddin in list)
            {
                bool found = false;
                foreach (string name in valueNames)
                {
                    if (itemAddin.ValueName == name)
                    {
                        found = true;
                        break;
                    }
                }
                if (!found)
                    deleteList.Add(itemAddin);
            }

            foreach (DisabledValue itemAddin in deleteList)
                list.Remove(itemAddin);
        }

        private static bool Contains(List<DisabledValue> list, string name)
        {
            foreach (DisabledValue value in list)
            {
                if (name == value.ValueName)
                    return true;
            }
            return false;
        }

        public DisabledValue[] Values
        {
            get
            {
                RegistryKey key = _rootKey.OpenSubKey(_registryPath);
                if (null != key)
                {
                    string[] names = key.GetValueNames();
                    foreach (string item in names)
                    {
                        if (!Contains(_valuesList, item))
                        {
                            object regValue = key.GetValue(item, null);
                            regValue = DisabledValue.ConvertDisabledItemValue(regValue);
                            DisabledValue newValue = new DisabledValue(_parent, this, item, regValue);
                            _valuesList.Add(newValue);
                        }
                    }
                    key.Close();
                    DeleteNotExistingItems(_valuesList, names);
                    return _valuesList.ToArray();
                }
                return new DisabledValue[0];
            }
        }

        public string Name
        {
            get
            {
                return _name;
            }
        }

        public string OfficeProductVersion
        {
            get
            {
                string[] splitArray = RegistryPath.Split(new string[] { "\\" }, StringSplitOptions.RemoveEmptyEntries);
                return splitArray[4] + " " + splitArray[3];
            }
        }

        public int ValueCount
        {
            get
            {
                RegistryKey key = _rootKey.OpenSubKey(_registryPath);
                if (null != key)
                {
                    int valueCount = key.ValueCount;
                    key.Close();
                    return valueCount;
                }
                return 0;
            }
        }

        internal NotifyKind CheckChangedValueCount(ref RegistryChangeInfo changeInfos)
        {
            RegistryKey key = _rootKey.OpenSubKey(_registryPath);
            if (null == key)
            {
                bool found = false;
                foreach (string keyExists in _keyExists)
                {
                    if (keyExists == (RootKey.ToString() + "\\" + _registryPath))
                    {
                        found = true;
                        break;
                    }
                }
                if (found)
                { 
                    _parent.RaisePropertyChanged(this);
                    _keyExists.Remove(RootKey.ToString() + "\\" + _registryPath);
                }
            }
            else 
            { 
                bool found = false;
                foreach (string keyExists in _keyExists)
                {
                    if(keyExists == (RootKey.ToString() +"\\" + _registryPath))
                    {
                        found = true;
                        break;
                    }
                }
                if(!found)
                    _keyExists.Add(RootKey.ToString() + "\\" + _registryPath);

                string[] valueNames = key.GetValueNames();
                int valueCount = key.ValueCount;
                if (valueCount != _lastValueCount)
                {
                    _parent.RaisePropertyChanged(this);

                    NotifyKind returnKind;
                    if ((valueCount > _lastValueCount) && (!_parent.FirstRun))
                    {
                        NewDeactivatedElementInfo newElementInfo = new NewDeactivatedElementInfo();
                        newElementInfo.RootKey = _rootKey;
                        newElementInfo.KeyPath = _registryPath;

                        string newValueName = GetNewValueName(valueNames, _valueNames);
                        object regValue = key.GetValue(newValueName, null);
                        regValue = DisabledValue.ConvertDisabledItemValueName(regValue, newValueName);
                        newElementInfo.Name = regValue as string;
                        newElementInfo.OfficeProductVersion = OfficeProductVersion;
                        changeInfos = newElementInfo;
                        _valueNames = valueNames;
                        returnKind = NotifyKind.DisabledItemNew;
                    }
                    else if (!_parent.FirstRun)
                    {
                        DeleteDeactivatedElementInfo deleteElementInfo = new DeleteDeactivatedElementInfo();
                        deleteElementInfo.RootKey = _rootKey;
                        deleteElementInfo.KeyPath = _registryPath;
                        deleteElementInfo.Name = GetDeletedValueName(valueNames, _valueNames);

                        string convertedName = null;
                        _convertedNames.TryGetValue(deleteElementInfo.Name, out convertedName);
                        _convertedNames.Remove(deleteElementInfo.Name);
                        deleteElementInfo.Name = convertedName;

                        deleteElementInfo.OfficeProductVersion = OfficeProductVersion;
                        changeInfos = deleteElementInfo;
                        _valueNames = valueNames;
                        returnKind = NotifyKind.DisabledItemDelete;
                    }
                    else
                    {
                        foreach (string name in valueNames)
                        {
                            string existingName = null;
                            _convertedNames.TryGetValue(name, out existingName);
                            if (null == existingName)
                            {
                                object regValue = key.GetValue(name, null);
                                regValue = DisabledValue.ConvertDisabledItemValueName(regValue, name);
                                _convertedNames.Add(name, regValue as string);
                            }
                        }
                        returnKind = NotifyKind.Nothing;
                    }


                    _lastValueCount = valueCount;
                    return returnKind;
                }
                _valueNames = valueNames;
                key.Close();
            }
            return NotifyKind.Nothing;
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

        public DisabledKey(WatchController parent, string name, RegistryKey rootKey, string registryPath)
        {
            _parent = parent;
            _rootKey = rootKey;
            _name = name;
            _registryPath = registryPath;
            GetKeyValueCount();
        }

        #endregion

        #region Private Methods

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

        private void GetKeyValueCount()
        {
            RegistryKey key = _rootKey.OpenSubKey(_registryPath);
            if (null != key)
            {
                _valueCount = key.ValueCount;
                key.Close();
            }
        }

        internal static bool IsEqual(object value1, object value2)
        { 
            if((null == value1) && (null == value2))
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
    }
}
