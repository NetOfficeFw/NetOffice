using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;

namespace NetOffice.DeveloperToolbox.Utils.Registry
{
    public class UtilsRegistryEntry
    {
        #region Fields

        private UtilsRegistryKey        _parent;
        private string                  _valueName;
        private UtilsRegistryEntryType  _type;

        #endregion

        #region Construction

        internal UtilsRegistryEntry(UtilsRegistryKey parent, string valueName)
        {
            _parent = parent;
            _valueName = valueName;
            _type = UtilsRegistryEntryType.Normal;
        }

        internal UtilsRegistryEntry(UtilsRegistryKey parent, string valueName, UtilsRegistryEntryType type)
        {
            _parent = parent;
            _valueName = valueName;
            _type = type;
        }

        #endregion

        #region Properties

        public UtilsRegistryEntryType Type
        {
            get
            {
                return _type;
            }
        }

        public string Name
        {
            get
            {
                if(string.IsNullOrEmpty(_valueName))
                    return "(Standard)";
                else
                    return _valueName;
            }
            set
            {
                RegistryKey key = _parent.Open(true);
                RegistryValueKind regKind = key.GetValueKind(_valueName);
                object regValue = key.GetValue(_valueName);
                key.DeleteValue(_valueName);
                key.SetValue(value, regValue, regKind);
                key.Close();
                _valueName = value;
            }
        }

        public object Value
        {
            get
            {
                if (_type == UtilsRegistryEntryType.Faked)
                    return null;

                RegistryKey key = _parent.Open();
                object regValue = key.GetValue(_valueName);
                key.Close();
                return regValue;
            }
            set
            {
                if (_type == UtilsRegistryEntryType.Faked)
                {
                    RegistryKey key = _parent.Open(true);
                    key.SetValue(_valueName, value, ValueKind);
                    key.Close();
                    _type = UtilsRegistryEntryType.Default; 
                }
                else
                {
                    RegistryKey key = _parent.Open(true);
                    key.SetValue(_valueName, value, ValueKind);
                    key.Close();
                }
            }
        }
        
        public RegistryValueKind ValueKind
        {
            get
            {
                if (_type == UtilsRegistryEntryType.Faked)
                    return RegistryValueKind.String;

                RegistryKey key = _parent.Open();
                RegistryValueKind kind = key.GetValueKind(_valueName);
                key.Close();
                return kind;
            }
        }       

        #endregion

        #region Methods

        public static string ByteArrayToBinaryString(byte[] byteArray)
        {

            StringBuilder builder = new StringBuilder(byteArray.Length * 2);
            foreach (byte value in byteArray)
            {
                builder.AppendFormat("{0:X2}", value);
                builder.Append(" ");
            }
            return builder.ToString();
        }

        public static string ShiftHexValue(string value)
        {
            int lenght = value.Length;
            if ((10 - lenght) >= 2)
                value = "0x" + value;
            lenght = value.Length;
            if ((10 - lenght) > 0)
            {
                for (int i = 0; i < (10 - lenght); i++)
                {
                    string first = value.Substring(0, 2);
                    string last = value.Substring(2);
                    value = first + "0" + last;
                }
            }
            return value;
        }

        public string GetValue(int lcid = 1033)
        {
            RegistryValueKind kind = ValueKind;
            switch (kind)
            {
                case RegistryValueKind.DWord:
                case RegistryValueKind.QWord:
                    return ShiftHexValue(String.Format("{0:x4}", Value)) + " (" + Convert.ToString(Value) + ")";
                case RegistryValueKind.Binary:
                    return ByteArrayToBinaryString((Value as byte[])).ToLower();
                case RegistryValueKind.ExpandString:
                case RegistryValueKind.MultiString:
                case RegistryValueKind.String:
                case RegistryValueKind.Unknown:
                    if (_type == UtilsRegistryEntryType.Faked)
                        return lcid == 1033 ? "(Value not set)" : "(Wert nicht gesetzt)";
                    else if ((_type == UtilsRegistryEntryType.Default) && (null == Value))
                        return lcid == 1033 ? "(Value not set)" : "(Wert nicht gesetzt)";
                    return Value as string;
                default:
                    throw new ArgumentException(kind.ToString() + " is out of range");
            }
        }

        public void Delete()
        {
            RegistryKey key = _parent.Open(true);
            key.DeleteValue(_valueName);
            key.Close();
        }

        #endregion

        #region Static Methods

        private static byte[] StringToByteArray(string str)
        {
            if (null == str)
                return null;
            System.Text.UnicodeEncoding enc = new System.Text.UnicodeEncoding();
            return enc.GetBytes(str);
        }

        private static string ByteArrayToString(byte[] arr)
        {
            if (null == arr)
                return null;
            System.Text.UnicodeEncoding enc = new System.Text.UnicodeEncoding();
            return enc.GetString(arr);
        }

        #endregion

        #region Overrides

        public override string ToString()
        {
            return String.Format("UtilsRegistryEntry {0}", Name);
        }

        #endregion
    }
}
