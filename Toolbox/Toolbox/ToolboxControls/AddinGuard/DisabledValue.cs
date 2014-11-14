using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.ToolboxControls.AddinGuard
{
    class DisabledValue
    {
        #region Fields

        WatchController _root;
        DisabledKey _parent;
        string _valueName;
        string _disabledItemName;
        object _value;

        #endregion

        #region Construction

        internal DisabledValue(WatchController root, DisabledKey item, string valueName, object value)
        {
            _root = root;
            _parent = item;
            _valueName = valueName;
            _disabledItemName = GetDisabledItemName(value as string);
            _value = value;
        }

        #endregion

        #region Properties

        public DisabledKey Parent
        {
            get
            {
                return _parent;
            }
        }

        public string ValueName
        {
            get
            {
                return _valueName;
            }
        }

        public string Name
        {
            get
            {
                return _disabledItemName;
            }
        }

        public string OfficeProductVersion
        {
            get
            {
                string[] splitArray = _parent.RegistryPath.Split(new string[] { "\\" }, StringSplitOptions.RemoveEmptyEntries);
                return splitArray[4] + " " + splitArray[3];
            }
        }

        public string Value
        {
            get
            {
                return _value as string;
            }
        }

        #endregion

        #region Methods

        internal string GetDisabledItemName(string value)
        {
            if (null == value)
                return _valueName;

            int i = value.LastIndexOf("\\", StringComparison.InvariantCultureIgnoreCase);
            return value.Substring(i + 1);
        }

        internal static string ConvertDisabledItemValueName(object value, string faultyName)
        {
            string name = ConvertDisabledItemValue(value);
            if (!string.IsNullOrEmpty(name))
            {
                int i = name.LastIndexOf("\\", StringComparison.InvariantCultureIgnoreCase);
                name = name.Substring(i + 1);
            }
            else
                name = faultyName;

            return name;
        }

        internal static string ConvertDisabledItemValue(object value)
        {
            if (null != value)
            {
                if (!(value is byte[]))
                    return null;
                byte[] byteArray = (byte[])value;
                if (byteArray.Length > 2)
                {
                    string val = ByteArrayToString(byteArray);
                    val = val.Replace("\0", "");
                    int i = val.LastIndexOf("\\", StringComparison.InvariantCultureIgnoreCase);
                    if (i >= 0)
                        return ValidateByteArrayString(val.Substring(i - 2));
                    else
                    {
                        return ValidateByteArrayString(val);
                    }

                }
                else
                    return null;
            }
            else
                return null;
        }

        internal static string ValidateByteArrayString(string value)
        {
            if (null == value)
                return value;

            string result = "";
            foreach (char item in value.ToCharArray())
            {
                if ("abcdefghijklmnopqrstuvwxyz,.-:;()*".IndexOf(item.ToString(), StringComparison.InvariantCultureIgnoreCase) >= 0)
                    result += item;
            }
            return result;
        }

        internal static string ByteArrayToString(byte[] arr)
        {
            System.Text.UnicodeEncoding enc = new System.Text.UnicodeEncoding();
            return enc.GetString(arr);
        }

        #endregion
    }
}
