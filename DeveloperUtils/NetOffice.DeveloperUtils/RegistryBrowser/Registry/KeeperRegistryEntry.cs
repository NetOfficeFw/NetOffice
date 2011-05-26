using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;

namespace NetOffice.DeveloperUtils.RegistryBrowser
{
    public class KeeperRegistryEntry
    {
        #region Member

        private RegistryKey        _root;
        private string             _key;

        private string             _valueName;
        private object             _value;
        private RegistryValueKind  _valueType;

        private bool               _isBinary;

        #endregion

        #region Properties

        public object Value
        {
            get
            {
                return _value;
            }
            set            
            {
                RegistryKey rk = _root.OpenSubKey(_key, true);

                if (_isBinary)
                {
                    byte[] arr = StringToByteArray((string)value);
                    rk.SetValue(_valueName, arr);
                    rk.Close();
                }
                else
                { 
                    rk.SetValue(_valueName, value);
                    rk.Close();
                }
                
                _value = value;
            }
        }

        public RegistryValueKind ValueType
        {
            get
            {
                return _valueType;
            }
        }
        public string Name
        {
            get 
            {
                return _valueName; 
            }
        }

        #endregion

        #region Methods

        public void Refresh()
        {
            RegistryKey rk = _root.OpenSubKey(_key, true);
            _valueType = rk.GetValueKind(_valueName);
            _value = rk.GetValue(_valueName);

            if (_value is byte[])
                _value = ByteArrayToString((byte[])_value);
            else
                _value = rk.GetValue(_valueName);
                
            rk.Close();
        }

        #endregion

        private byte[] StringToByteArray(string str)
        {
            System.Text.ASCIIEncoding enc = new System.Text.ASCIIEncoding();
            return enc.GetBytes(str);
        }

        private static string ByteArrayToString(byte[] arr)
        {
            System.Text.ASCIIEncoding enc = new System.Text.ASCIIEncoding();
            return enc.GetString(arr);
        }

        #region Construction

        internal KeeperRegistryEntry(RegistryKey root, string key, string ValueName, object Value, RegistryValueKind ValueType, bool isBinary)
        {
            _root       = root;
            _key        = key;
            _valueName  = ValueName;

            _value      = Value;
            _valueType  = ValueType;

            _isBinary = isBinary;
        }

        #endregion
    }
}
