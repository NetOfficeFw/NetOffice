using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddin.RegFile
{
    internal static class RegValueConverter
    {
        internal static string ToDwordString(object value)
        {
            try
            {
                if (null != value)
                    return ToDwordString((int)value);
                else
                    return ToDwordString(0);
            }
            catch
            {
                return ToDwordString(0);
            }
        }

        internal static string ToDwordString(int value)
        {
            string result = value.ToString();
            int numericLength = result.Length;
            int paddingLeftCount = 8 - result.Length;
            result = result.PadLeft(paddingLeftCount + numericLength, '0');
            return result;
        }

        internal static string EncryptExpandString(string text)
        {
            return EncryptExpandString(text, 13);
        }

        internal static string EncryptExpandString(string text, int lineLength)
        {
            StringBuilder result = new StringBuilder();
            result.Append("hex(2):");
            List<string> hex = EncryptString(text);
            int lineCharIndex = 0;

            for (int i = 0; i < hex.Count; i++)
            {
                string item = hex[i];
                result.Append(item + ",00,");

                lineCharIndex++;
                if (lineCharIndex > lineLength)
                {
                    result.Append("\\" + Environment.NewLine + "  ");
                    lineCharIndex = 0;
                }
            }
            result.Append("00,00");
            return result.ToString();
        }

        internal static string EncryptMultiString(string[] text)
        {
            StringBuilder result = new StringBuilder();
            List<string> list = new List<string>();
            result.Append("hex(7):");
            foreach (string item in text)
            {
                List<string> hexList = EncryptString(item);

                foreach (var hexItem in hexList)
                    list.Add(hexItem + ",00");

                list.Add("00");
                list.Add("00");
            }

            list.Add("00");
            list.Add("00");

            for (int i = 0; i < list.Count; i++)
            {
                string item = list[i];
                if (i == list.Count - 1)
                    result.Append(item);
                else
                    result.Append(item + ",");
            }

            return result.ToString();
        }

        internal static string EncryptBinary(byte[] value)
        {
            StringBuilder result = new StringBuilder();
            result.Append("hex:");
            List<string> hex = EncryptBytes(value);
            for (int i = 0; i < hex.Count; i++)
            {
                string item = hex[i];
                if (i == hex.Count - 1)
                    result.Append(item);
                else
                    result.Append(item + ",");
            }

            return result.ToString();
        }

        internal static string EncryptQ(Int64 value)
        {
            string result = "";
            byte[] temp = BitConverter.GetBytes(value);
            temp.Reverse();
            foreach (byte item in temp)
            {
                string hex = String.Format("{0:X}", item).ToLower();
                if (hex.Length == 1)
                    hex = hex.PadLeft(2, '0');
                result += hex;
            }

            result.Reverse();

            string split1 = result.Substring(0, 2);
            string split2 = result.Substring(2, 2);
            string split3 = result.Substring(4, 2);
            string split4 = result.Substring(6, 2);
            string split5 = result.Substring(8, 2);
            string split6 = result.Substring(10, 2);
            string split7 = result.Substring(12, 2);
            string split8 = result.Substring(14, 2);

            return String.Format("hex(b):{0},{1},{2},{3},{4},{5},{6},{7}", split1, split2, split3, split4, split5, split6, split7, split8);
        }

        private static List<string> EncryptBytes(byte[] values)
        {
            List<string> list = new List<string>();
            foreach (byte letter in values)
            {
                int value = Convert.ToInt32(letter);
                string hex = String.Format("{0:X}", value).ToLower();
                hex = hex.PadLeft(2, '0');
                list.Add(hex);
            }
            return list;
        }

        private static List<string> EncryptString(string text)
        {
            List<string> list = new List<string>();
            char[] values = text.ToCharArray();
            foreach (char letter in values)
            {
                int value = Convert.ToInt32(letter);
                string hex = String.Format("{0:X}", value).ToLower();
                hex = hex.PadLeft(2, '0');
                list.Add(hex);
            }
            return list;
        }

    }
}
