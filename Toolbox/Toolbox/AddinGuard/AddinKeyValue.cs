using System;
using System.Collections.Generic;
using Microsoft.Win32;
using System.Text;
using NetOffice.DeveloperToolbox.WindowsRegistry;

namespace NetOffice.DeveloperToolbox.AddinGuard
{
    class AddinKeyValue
    {
        #region Fields

        WatchController _root;
        AddinKey _parent;
        string _valueName;
        RegistryValueKind _valueKind;
        object _value;
        #endregion

        #region Construction

        internal AddinKeyValue(WatchController root, AddinKey parent, string valueName, RegistryValueKind valueKind, object value)
        {
            _root = root;
            _parent = parent;
            _valueName = valueName;
            _valueKind = valueKind;
            _value = value;
        }

        #endregion

        #region Properties

        public AddinKey Parent
        {
            get
            {
                return _parent;
            }
        }

        public string Name
        {
            get 
            {
                return _valueName;
            }
        }

        public RegistryValueKind Type
        {
            get 
            {
                return _valueKind; 
            }
        }

        public object Value
        {
            get
            {
                if (RegistryValueKind.Binary == _valueKind)
                {
                    string val = UtilsRegistryEntry.ByteArrayToBinaryString((_value as byte[]));
                    if(null != val)
                        val = val.ToLower();
                    return val;
                }
                else
                    return _value;
            }
        }

        #endregion
    }
}
