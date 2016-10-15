using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;   
using System.Linq;
using System.Reflection;
using Microsoft.Win32;

namespace RegAddin.Common
{
    internal class AddinRegAnalyzer
    {
        public static string[] _multiRegisterIn = new string[] { "Excel", "Word", "Outlook", "PowerPoint", "Access", "MS Project", "Visio" };

        private static string _systemObject = "System.Object";

        private static string _addinBaseClassName = "NetOffice.Tools.COMAddinBase";

        private static string[] _regLocationAttributeName = new string[] { "NetOffice.Tools.RegistryLocationAttribute" };

        private static string[] _multiRegisterName = new string[] { "NetOffice.OfficeApi.Tools.MultiRegisterAttribute" };

        private static string[] _classNames = new string[] { "NetOffice.MSProjectApi.Tools.COMAddin",
                                                        "NetOffice.ExcelApi.Tools.COMAddin",
                                                        "NetOffice.WordApi.Tools.COMAddin",
                                                        "NetOffice.OutlookApi.Tools.COMAddin",
                                                        "NetOffice.PowerPointApi.Tools.COMAddin",
                                                        "NetOffice.AccessApi.Tools.COMAddin",
                                                        "NetOffice.VisioApi.Tools.COMAddin",
                                                        "NetOffice.OfficeApi.Tools.COMAddin"};

        private static string[] _classKeys = new string[] { "MS Project", "Excel", "Word", "Outlook", "PowerPoint", "Access", "Visio"};

        private static string _multiClassName = "NetOffice.OfficeApi.Tools.COMAddin";

        private static string _attributeName = "NetOffice.Tools.COMAddinAttribute";

        private static string _officeRelatedKey = "Software\\Microsoft\\Office\\{0}\\Addins";

        internal void DeleteKey(Type addin, IEnumerable<object> addinAttributes, KeyTarget keyTarget)
        {
            Dictionary<object, Type> attributeTypes = GetAttributeTypes(addinAttributes);
            KeyValuePair<object, Type> progId = GetAttribute<ProgIdAttribute>(attributeTypes);
            KeyValuePair<object, Type> guid = GetAttribute<GuidAttribute>(attributeTypes);
            KeyValuePair<object, Type> comAddin = GetAttribute(attributeTypes, _attributeName);
            if (comAddin.Value == null)
                return;
            KeyValuePair<object, Type> reg = GetAttribute(attributeTypes, _regLocationAttributeName);
            KeyValuePair<object, Type> multi = GetAttribute(attributeTypes, _multiRegisterName);
            bool isMultiAddin = IsMultiAddin(addin);
            if (isMultiAddin && multi.Key == null)
                return;

            string progIdValue = (progId.Key as ProgIdAttribute).Value;

            int regLocation = 0;
            if (null != reg.Key)
                regLocation = (int)reg.Value.InvokeMember("Value", BindingFlags.Instance | BindingFlags.Public | BindingFlags.GetField, null, reg.Key, new string[0]);
            if (regLocation < 0 || regLocation > 2)
                throw new ArgumentException("regLocation");
            if (1 == regLocation)
                keyTarget = KeyTarget.User;
            else if (2 == regLocation)
                keyTarget = KeyTarget.System;

            if (isMultiAddin)
            {
                IEnumerable products = multi.Value.InvokeMember("Products", BindingFlags.Instance | BindingFlags.Public | BindingFlags.GetField, null, multi.Key, new string[0]) as IEnumerable;
                if (null != products)
                {
                    foreach (object item in products)
                    {
                        int productIndex = Convert.ToInt32(item);
                        DeleteOfficeRegistryKey(_multiRegisterIn[productIndex], progIdValue, keyTarget);
                    }
                }
            }
            else
            {
                string key = GetKeyName(addin);
                DeleteOfficeRegistryKey(key, progIdValue, keyTarget);
            }
        }

        internal void CreateKey(Type addin, IEnumerable<object> addinAttributes, bool useSystemKey)
        {
            Dictionary<object, Type> attributeTypes = GetAttributeTypes(addinAttributes);
            KeyValuePair<object, Type> progId = GetAttribute<ProgIdAttribute>(attributeTypes);
            KeyValuePair<object, Type> guid = GetAttribute<GuidAttribute>(attributeTypes);
            KeyValuePair<object, Type> comAddin = GetAttribute(attributeTypes, _attributeName);
            if (comAddin.Value == null)
                return;
            KeyValuePair<object, Type> reg = GetAttribute(attributeTypes, _regLocationAttributeName);
            KeyValuePair<object, Type> multi = GetMultiRegisterAttribute(attributeTypes);
            bool isMultiAddin = IsMultiAddin(addin);
            if (isMultiAddin && multi.Key == null)
                return;

            string progIdValue = (progId.Key as ProgIdAttribute).Value;

            int regLocation = 0;           
            if (null != reg.Key)
                regLocation = (int)reg.Value.InvokeMember("Value", BindingFlags.Instance | BindingFlags.Public | BindingFlags.GetField, null, reg.Key, new string[0]);
            if (regLocation < 0 || regLocation > 2)
                throw new ArgumentException("regLocation");
            if (0 == regLocation)
                useSystemKey = regLocation == 2 ? true : false;

            string name = (string)comAddin.Value.InvokeMember("Name", BindingFlags.Instance |BindingFlags.Public | BindingFlags.GetField, null, comAddin.Key, new object[0]);
            string description = (string)comAddin.Value.InvokeMember("Description", BindingFlags.Instance | BindingFlags.Public | BindingFlags.GetField, null, comAddin.Key, new object[0]);
            int loadBehavior = (int)comAddin.Value.InvokeMember("LoadBehavior", BindingFlags.Instance | BindingFlags.Public | BindingFlags.GetField, null, comAddin.Key, new object[0]);
            int commandLineSafe = (int)comAddin.Value.InvokeMember("CommandLineSafe", BindingFlags.Instance | BindingFlags.Public | BindingFlags.GetField, null, comAddin.Key, new object[0]);

            if (isMultiAddin)
            {
                IEnumerable products = multi.Value.InvokeMember("Products", BindingFlags.Instance | BindingFlags.Public | BindingFlags.GetField, null, multi.Key, new string[0]) as IEnumerable;
                if (null != products)
                {
                    foreach (object item in products)
                    {
                        int productIndex = Convert.ToInt32(item);
                        CreateOfficeRegistryKey(_multiRegisterIn[productIndex], progIdValue, name, description, loadBehavior, commandLineSafe, useSystemKey);
                    }
                }
            }
            else
            {
                string key = GetKeyName(addin);
                CreateOfficeRegistryKey(key, progIdValue, name, description, loadBehavior, commandLineSafe, useSystemKey);
            }
        }

        internal enum KeyTarget
        {
            System = 0,
            User = 1,
            Both = 2
        }

        private void DeleteOfficeRegistryKey(string officeKeyName, string addinProgId, KeyTarget target)
        {
            RegistryKey[] rootKey = null;

            switch (target)
            {
                case KeyTarget.System:
                    rootKey = new RegistryKey[] { RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Default) };
                    break;
                case KeyTarget.User:
                    rootKey = new RegistryKey[] { RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default) };
                    break;
                case KeyTarget.Both:
                    rootKey = new RegistryKey[] { RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Default),
                                                  RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default) };
                    break;
                default:
                    throw new IndexOutOfRangeException("target");
            }

            foreach (var item in rootKey)
            {
                string targetKey = String.Format(_officeRelatedKey + "\\{1}", officeKeyName, addinProgId);
                try
                {
                    item.DeleteSubKeyTree(targetKey, false);
                    item.Close();
                }
                catch (System.Security.SecurityException)
                {
                    ;
                }
                catch (System.UnauthorizedAccessException)
                {
                    ;
                }
                catch (Exception)
                {
                    throw;
                }
             
            }
        }

        private void CreateOfficeRegistryKey(string officeKeyName, string addinProgId, string name, string description, int loadBehavior, int commandLineSafe, bool useSystemKey)
        {
            RegistryKey rootKey = null;
            if(useSystemKey)
                rootKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Default);
            else
                rootKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default);

            string targetKey = String.Format(_officeRelatedKey + "\\{1}", officeKeyName, addinProgId);

            RegistryKey officeKey = rootKey.CreateSubKey(targetKey);

            officeKey.SetValue("LoadBehavior", loadBehavior, RegistryValueKind.DWord);
            officeKey.SetValue("FriendlyName", name, RegistryValueKind.String);
            officeKey.SetValue("Description", description, RegistryValueKind.String);       
            if(commandLineSafe != -1)
                officeKey.SetValue("CommandLineSafe", commandLineSafe, RegistryValueKind.DWord);
            officeKey.Close();
            rootKey.Close();
        }
         
        private string GetKeyName(Type addin)
        {
            Type target = addin;
            while (null != target)
            {
                if (target.BaseType.FullName == _systemObject || target.BaseType.FullName == _addinBaseClassName)
                    break;
                target = target.BaseType;
            }

            int index = -1;
            for (int i = 0; i < _classNames.Length; i++)
            {
                if (_classNames[i] == target.FullName)
                {
                    index = i;
                    break;
                }
            }            
            return _classKeys[index];
        }

        internal static KeyValuePair<object, Type> GetMultiRegisterAttribute(Dictionary<object, Type> attributeTypes)
        {
            string name = _multiRegisterName[0];
            foreach (KeyValuePair<object, Type> item in attributeTypes)
            {
                if (name == item.Value.FullName)
                    return item;
            }
            return default(KeyValuePair<object, Type>);
        }

        internal static KeyValuePair<object, Type> GetComAddinAttribute(Dictionary<object, Type> attributeTypes)
        {
            string name = _attributeName;
            foreach (KeyValuePair<object, Type> item in attributeTypes)
            {
                if (name == item.Value.FullName)
                    return item;
            }
            return default(KeyValuePair<object, Type>);
        }

        private KeyValuePair<object, Type> GetAttribute(Dictionary<object, Type> attributeTypes, string name)
        {
            foreach (KeyValuePair<object, Type> item in attributeTypes)
            {
                if (name == item.Value.FullName)
                    return item;
            }
            return default(KeyValuePair<object, Type>);
        }

        private KeyValuePair<object, Type> GetAttribute(Dictionary<object, Type> attributeTypes, string[] name)
        {
            foreach (KeyValuePair<object, Type> item in attributeTypes)
            {                
                if (name.Contains(item.Value.FullName))
                    return item;
            }
            return default(KeyValuePair<object, Type>);
        }

        private KeyValuePair<object,Type> GetAttribute<T>(Dictionary<object, Type> attributeTypes)
        {
            foreach (KeyValuePair<object, Type> item in attributeTypes)
            {
                if (item.Key is T)
                    return item;
            }
            throw new ArgumentOutOfRangeException();
        }

        private static Dictionary<object, Type> GetAttributeTypes(IEnumerable<object> attributes)
        {
            Dictionary<object, Type> result = new Dictionary<object, Type>();
            foreach (object item in attributes)
                result.Add(item, item.GetType());
            return result;
        }

        internal static bool IsMultiAddin(Type addin)
        {
            Type target = addin;
            while (null != target)
            {
                if (target.BaseType.FullName == _systemObject || target.BaseType.FullName == _addinBaseClassName)
                    break;
                target = target.BaseType;
            }

            return target.FullName == _multiClassName;
        }
    }
}
