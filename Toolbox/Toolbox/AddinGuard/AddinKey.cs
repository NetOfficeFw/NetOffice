using System;
using System.Collections.Generic;
using Microsoft.Win32;
using System.Text;

namespace NetOffice.DeveloperToolbox.AddinGuard
{
    class AddinKey
    {
        #region Fields

        WatchController _root;
        AddinsKey _parent;
        string _registryPath;
        string _subKeyName;

        string[] _valueNames;
        Dictionary<string, RegistryValueKind> _valueKinds;
        Dictionary<string, object> _values;

        int _valueCount;

        #endregion

        #region Construction

        internal AddinKey(WatchController root, AddinsKey item, string registryPath, string subKeyName)
        {
            _root = root;
            _parent = item;
            _registryPath = registryPath;
            _subKeyName = subKeyName;
            GetKeyValueCount();
            GetValueInfos();
        }

        #endregion

        #region Properties

        public AddinKeyValue[] Values
        {
            get
            {
                RegistryKey key = _parent.RootKey.OpenSubKey(_registryPath);
                if (null != key)
                {
                    List<AddinKeyValue> list = new List<AddinKeyValue>();
                    string[] names = key.GetValueNames();
                    foreach (string item in names)
                    {

                        AddinKeyValue newValue = new AddinKeyValue(_root, this, item, key.GetValueKind(item), key.GetValue(item, null));
                        list.Add(newValue);
                    }
                    key.Close();
                    return list.ToArray();
                }
                return new AddinKeyValue[0];
            }
        }

        public string Name
        {
            get
            {
                return _subKeyName;
            }
        }

        public AddinsKey Parent
        {
            get
            {
                return _parent;
            }
        }

        public int? LoadBehavior
        {
            get
            {
                RegistryKey key = _parent.RootKey.OpenSubKey(_registryPath);
                if (null != key)
                {
                    int? value = key.GetValue("LoadBehavior", null) as int?;
                    key.Close();
                    return value;
                }
                return null;
            }
        }

        public int ValueCount
        {
            get
            {
                RegistryKey key = _parent.RootKey.OpenSubKey(_registryPath);
                if (null != key)
                {
                    int valueCount = key.ValueCount;
                    key.Close();
                    return valueCount;
                }
                return 0;
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

        #region Methods

        internal NotifyKind CheckChangedValues(ref RegistryChangeInfo changeInfos)
        {
            try
            {
                bool openModeWrite = true;
                if ((_parent.RootKey == Registry.LocalMachine) && (_root.ReadOnlyModeForMachineKeys))
                    openModeWrite = false;

                RegistryKey key = _parent.RootKey.OpenSubKey(_registryPath, openModeWrite);
                if (null != key)
                {
                    string[] valueNames = key.GetValueNames();
                    int valueCount = key.ValueCount;
                    if (valueCount != _valueCount)
                    {
                        _root.RaisePropertyChanged(Parent);
                        NotifyKind returnKind;
                        if (valueCount > _valueCount)
                        {
                            AddinValuesIncrementInfo incrementInfo = new AddinValuesIncrementInfo();
                            incrementInfo.RootKey = _parent.RootKey;
                            incrementInfo.KeyPath = _registryPath;
                            incrementInfo.KeyName = _parent.Name;
                            incrementInfo.ValueName = GetNewValueName(valueNames, _valueNames);
                            _valueNames = valueNames;
                            _valueKinds.Add(incrementInfo.ValueName, key.GetValueKind(incrementInfo.ValueName));
                            _values.Add(incrementInfo.ValueName, key.GetValue(incrementInfo.ValueName));
                            changeInfos = incrementInfo;
                            returnKind = NotifyKind.AddinValuesIncrement;
                        }
                        else
                        {
                            AddinValuesDecrementInfo decrementInfo = new AddinValuesDecrementInfo();
                            decrementInfo.RootKey = _parent.RootKey;
                            decrementInfo.KeyPath = _registryPath;
                            decrementInfo.KeyName = _parent.Name;
                            decrementInfo.ValueName = GetDeletedValueName(valueNames, _valueNames);
                            _valueKinds.Remove(decrementInfo.ValueName);
                            _values.Remove(decrementInfo.ValueName);
                            changeInfos = decrementInfo;
                            returnKind = NotifyKind.AddinValuesDecrement;
                        }
                        _valueNames = valueNames;
                        _valueCount = valueCount;
                        key.Close();
                        return returnKind;
                    }
                    else
                    {
                        foreach (string item in valueNames)
                        {
                            // name changed
                            if (!AddinsKey.Contains(_valueNames, item))
                            {
                                _root.RaisePropertyChanged(_parent);

                                key.Close();
                                AddinValueNameChangedInfo nameInfo = new AddinValueNameChangedInfo();
                                nameInfo.RootKey = _parent.RootKey;
                                nameInfo.KeyPath = _registryPath;
                                nameInfo.KeyName = _parent.Name;
                                nameInfo.NewValueName = item;
                                nameInfo.OldValueName = GetChangedValueName(valueNames, _valueNames);
                                _valueNames = valueNames;
                                RegistryValueKind refreshKind;
                                object refreshValue = null;
                                _valueKinds.TryGetValue(nameInfo.OldValueName, out refreshKind);
                                _values.TryGetValue(nameInfo.OldValueName, out refreshValue);
                                _valueKinds.Remove(nameInfo.OldValueName);
                                _values.Remove(nameInfo.OldValueName);
                                _valueKinds.Add(nameInfo.NewValueName, refreshKind);
                                _values.Add(nameInfo.NewValueName, refreshValue);
                                changeInfos = nameInfo;

                                return NotifyKind.AddinValueNameIsChanged;
                            }

                            // value changed
                            object value = key.GetValue(item, null);
                            object oldValue = null;
                            _values.TryGetValue(item, out oldValue);
                            if (!AddinsKey.IsEqual(value, oldValue))
                            {

                                if ((_root.RestoreLastLoadBehavior) && (item == "LoadBehavior") && (!_root.FirstRun) && (true == openModeWrite) && (true == IsRestoreSituation(oldValue, value)))                                
                                {
                                    key.SetValue("LoadBehavior", oldValue);
                                    key.Close();
                                    AddinValueValueRestoredInfo restoredInfo = new AddinValueValueRestoredInfo();
                                    restoredInfo.RootKey = _parent.RootKey;
                                    restoredInfo.KeyPath = _registryPath;
                                    restoredInfo.KeyName = _parent.Name;
                                    restoredInfo.ValueName = item;
                                    restoredInfo.OldValue = value;
                                    restoredInfo.RestoredValue = oldValue;
                                    changeInfos = restoredInfo;
                                    return NotifyKind.AddinLoadBehaviorRestored;
                                }
                                else
                                {
                                    _root.RaisePropertyChanged(_parent);
                                    _values[item] = value;
                                    key.Close();
                                    AddinValueValueChangedInfo valueInfo = new AddinValueValueChangedInfo();
                                    valueInfo.RootKey = _parent.RootKey;
                                    valueInfo.KeyPath = _registryPath;
                                    valueInfo.KeyName = _parent.Name;
                                    valueInfo.ValueName = item;
                                    valueInfo.NewValue = value;
                                    valueInfo.OldValue = oldValue;
                                    changeInfos = valueInfo;
                                    return NotifyKind.AddinValueIsChanged;
                                }
                            }

                            // kind changed
                            RegistryValueKind kind = key.GetValueKind(item);
                            RegistryValueKind oldkind;
                            _valueKinds.TryGetValue(item, out oldkind);
                            if (!AddinsKey.IsEqual(kind, oldkind))
                            {
                                _root.RaisePropertyChanged(_parent);
                                _valueKinds[item] = kind;
                                key.Close();

                                AddinValueKindChangedInfo kindInfo = new AddinValueKindChangedInfo();
                                kindInfo.RootKey = _parent.RootKey;
                                kindInfo.KeyPath = _registryPath;
                                kindInfo.KeyName = _parent.Name;
                                kindInfo.ValueName = item;
                                kindInfo.NewValueKind = kind;
                                kindInfo.OldValueKind = oldkind;
                                changeInfos = kindInfo;
                                return NotifyKind.AddinValueKindIsChanged;
                            }
                        }
                        key.Close();
                    }

                }
                return NotifyKind.Nothing;
            }
            catch (System.Security.SecurityException exception)
            {
                throw new Exception("", exception);
            }
        }

        private static bool IsRestoreSituation(object oldValue, object newValue)
        {
            try
            {
                if ((null == oldValue) || (null == newValue))
                    return false;
                int oldVal = Convert.ToInt32(oldValue);
                int newVal = Convert.ToInt32(newValue);
                return ((newVal == 2) && (oldVal == 3));
            }
            catch
            {
                return false;
            }
           
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

        private void GetKeyValueCount()
        {
            RegistryKey key = _parent.RootKey.OpenSubKey(_registryPath);
            if (null != key)
            {
                _valueCount = key.ValueCount;
                key.Close();
            }
        }

        private void GetValueInfos()
        {
            RegistryKey key = _parent.RootKey.OpenSubKey(_registryPath);
            if (null != key)
            {
                _valueNames = key.GetValueNames();
                _valueKinds = new Dictionary<string, RegistryValueKind>();
                _values = new Dictionary<string, object>();
                foreach (string name in _valueNames)
                {
                    _valueKinds.Add(name, key.GetValueKind(name));
                    _values.Add(name, key.GetValue(name, null));
                }
                key.Close();
            }
        }

        #endregion
    }
}
