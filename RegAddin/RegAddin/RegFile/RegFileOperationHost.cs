using System;
using System.Reflection;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.ComponentModel;
using System.Text;
using System.Runtime.InteropServices;
using RegAddin.Common;
using Microsoft.Win32;

namespace RegAddin.RegFile
{
    [Serializable]
    internal class RegFileOperationHost : MarshalByRefObject, Common.IAppDomainMethod
    {
        #region Fields

        private static string _regGeneration = "Windows Registry Editor Version 5.00";
        private static string _regExportAttributeName = "NetOffice.Tools.ComRegExportCallAttribute";

        #endregion

        #region Ctor

        public RegFileOperationHost()
        {
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
        }

        #endregion

        #region Properties

        internal RegFileOperationHostSettings Settings { get; private set; }

        private string AssemblyPath
        {
            get
            {
                return Path.GetDirectoryName(Settings.AssemblyPath);
            }
        }

        #endregion

        #region IAppDomainMethod

        void Common.IAppDomainMethod.SetConfig(object configInstance)
        {
            RegFileOperationHostSettings settings = configInstance as RegFileOperationHostSettings;
            if (null != settings)
                Settings = settings;
            else
                throw new ArgumentException("Invalid configuration type.");
        }

        int Common.IAppDomainMethod.ExecuteInDomain()
        {
            bool regFilePathCheckResult = PathUtils.TryCheckForValidLocalAndAbsoluteFileSystemPath(Settings.RegFilePath);
            if (!regFilePathCheckResult)
                return (int)ResultCodes.InvalidRegfilePath;

            AppDomain domain = AppDomain.CurrentDomain;
            Assembly addinAssembly = Assembly.LoadFile(Settings.AssemblyPath);
            IEnumerable<object> assemblyAttributes = AssemblyReflection.GetCustomAssemblyAttributes(addinAssembly);

            if (!AssemblyReflection.AssemblyIsComVisible(addinAssembly, assemblyAttributes))
                return (int)ResultCodes.AssemblyNotComVisible;

            List<string> results = new List<string>();

            Type[] types = addinAssembly.GetExportedTypes();
            foreach (Type item in types)
            {
                if (!item.IsClass)
                    continue;

                IEnumerable<object> addinClassAttributes = Common.AttributeReflection.GetCustomClassAttributes(item);
                if (!AddinClassReflection.IsValidAddinClass(addinClassAttributes, item.Attributes))
                    continue;

                string regContent = CreateRegistryFileContent(addinAssembly, assemblyAttributes, Settings.Mode, item, addinClassAttributes);
 
                object exportResult = DoExportCall(item);
                string exportValues = ProceedExportResult(exportResult);
                if (null != exportValues)
                    regContent += exportValues;

                results.Add(regContent);
            }

            if (results.Count == 0)
                return (int)ResultCodes.NothingFound;

            WriteRegistryContentToLocalFileSystem(results);

            return (int)ResultCodes.Okay;
        }

        #endregion

        private string ProceedExportResult(object reg)
        {
            StringBuilder result = new StringBuilder();

            bool useSystem = Settings.Mode == SingletonSettings.RegisterMode.System;
            string rootKey = useSystem ? "[HKEY_LOCAL_MACHINE\\" : "[HKEY_CURRENT_USER\\";

            bool firstFlag = true;
            IEnumerable list = reg as IEnumerable;
            if (null != list)
            {
                foreach (object item in list)
                {
                    Type itemType = item.GetType();

                    if (firstFlag)
                        firstFlag = false;
                    else
                        result.AppendLine(String.Empty);

                    string key = itemType.InvokeMember("Key",
                        BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                        null, item, new object[0]) as string;
                    if (null != key)
                    {
                        if (key.StartsWith("\\"))
                            key = key.Substring("\\".Length);
                        result.AppendLine(String.Format("{0}{1}]", rootKey, key));
                    }

                    IEnumerable values = itemType.InvokeMember("Value",
                        BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                        null, item, new object[0]) as IEnumerable;

                    if (values != null)
                    {
                        foreach (object value in values)
                        {
                            Type valueType = value.GetType();
                            string name = valueType.InvokeMember("Name", BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance, null, value, new object[0]) as string;
                            object val = valueType.InvokeMember("Value", BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance, null, value, new object[0]);
                            RegistryValueKind kind = (RegistryValueKind)valueType.InvokeMember("Kind", BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance, null, value, new object[0]);
                            if (String.IsNullOrWhiteSpace(name))
                                name = "@";
                            else
                                name = "\"" + name + "\"";

                            switch (kind)
                            {
                                case RegistryValueKind.String:
                                    result.AppendLine(String.Format("{0}=\"{1}\"", name, null != val ? val.ToString() : String.Empty));
                                    break;
                                case RegistryValueKind.ExpandString:                                    
                                    result.AppendLine(String.Format("{0}={1}", name, null != val ? RegValueConverter.EncryptExpandString(val.ToString(), 13) : String.Empty));
                                    break;
                                case RegistryValueKind.MultiString:
                                    if(val is string[])
                                        result.AppendLine(String.Format("{0}={1}", name, null != val ? RegValueConverter.EncryptMultiString(val as string[]) : String.Empty)); break;
                                case RegistryValueKind.Binary:
                                    if (val is byte[])
                                        result.AppendLine(String.Format("{0}={1}", name, null != val ? RegValueConverter.EncryptBinary(val as byte[]) : String.Empty)); break;
                                case RegistryValueKind.DWord:
                                    result.AppendLine(String.Format("{0}=dword:{1}", name, RegValueConverter.ToDwordString(val)));
                                    break;
                                case RegistryValueKind.QWord:
                                    result.AppendLine(String.Format("{0}={1}", name, RegValueConverter.EncryptQ(Convert.ToInt64(val))));
                                    break;
                                case RegistryValueKind.Unknown:
                                case RegistryValueKind.None:
                                default:
                                    break;
                            }
                        }
                    }
                }
            }

            if (result.Length == 0)
                return null;
            else
            {                  
                return Environment.NewLine + "; --- Custom Export ---" + Environment.NewLine + result.ToString();
            }
        }

        private object DoExportCall(Type addin)
        {
            if (Settings.ExportCall == SingletonSettings.RegExportCall.On)
            {
                IEnumerable<MethodInfo> methods = Dispatcher.MethodUtils.GetMethodsFromAddinBaseClass(addin, BindingFlags.NonPublic | BindingFlags.Static);
                foreach (MethodInfo method in methods)
                {
                    if (Dispatcher.MethodUtils.HasAttribute(method, _regExportAttributeName) &&
                         method.GetParameters().Length == 3)
                    {
                        int scope = Settings.Mode == SingletonSettings.RegisterMode.System ? 0 : 1;
                        int keyState = Settings.AddinRegMode == SingletonSettings.AddinRegMode.Off ? 0 : 1;                    
                        return Dispatcher.MethodUtils.CallMethodWithArgumentsAndReturnValue(method, addin, scope, keyState);
                    }
                }
            }
            return null;
        }
         
        private void WriteRegistryContentToLocalFileSystem(List<string> contentTable)
        {
            if (null == contentTable || contentTable.Count == 0)
                return;

            StringBuilder fullContent = new StringBuilder();
            fullContent.AppendLine(_regGeneration);
            fullContent.AppendLine(String.Empty);

            foreach (string item in contentTable)
            {
                fullContent.AppendLine(String.Format("; ---  Begin Addin Connect Class ---"));
                fullContent.AppendLine(String.Empty);
                fullContent.Append(item);
                fullContent.AppendLine(String.Empty);
                fullContent.AppendLine(String.Format("; ---  End Addin Connect Class ---"));
                fullContent.AppendLine(String.Empty);
            }

            if (File.Exists(Settings.RegFilePath))
                File.Delete(Settings.RegFilePath);
            File.AppendAllText(Settings.RegFilePath, fullContent.ToString(), Encoding.Unicode);            
        }

        public static string[] _multiRegisterIn = new string[] { "Excel", "Word", "Outlook", "PowerPoint", "Access", "MS Project", "Visio" };

        private string CreateRegistryFileContent(Assembly addinAssembly, IEnumerable<object> assemblyAttributes, SingletonSettings.RegisterMode mode,
            Type addinClassType, IEnumerable<object> addinClassAttributes)
        {
            AddinClassInformations addinClass = AddinClassInformations.Create(
                            addinAssembly, assemblyAttributes, mode, addinClassType, addinClassAttributes);

            StringBuilder content = new StringBuilder();

            content.AppendLine(String.Format("[{0}\\{1}]", addinClass.ClassesRoot, addinClass.ProgId));
            content.AppendLine(String.Format("@=\"{0}\"", addinClass.FullClassName));
            content.AppendLine(String.Empty);

            content.AppendLine(String.Format("[{0}\\{1}\\CLSID]", addinClass.ClassesRoot, addinClass.ProgId));
            content.AppendLine("@=\"{" + addinClass.Id.ToString() + "}\"");
            content.AppendLine(String.Empty);

            content.AppendLine("[" + addinClass.ClassesRoot + "\\CLSID\\{" + addinClass.Id + "}]");
            content.AppendLine(String.Format("@=\"{0}\"", addinClass.FullClassName));
            content.AppendLine(String.Empty);

            content.AppendLine("[" + addinClass.ClassesRoot + "\\CLSID\\{" + addinClass.Id + "}\\InprocServer32]");
            content.AppendLine("@=\"mscoree.dll\"");
            content.AppendLine("\"ThreadingModel\"=\"Both\"");
            content.AppendLine(String.Format("\"Class\"=\"{0}\"", addinClass.FullClassName));
            content.AppendLine(String.Format("\"Assembly\"=\"{0}, Version={1}, Culture={2}, PublicKeyToken={3}\"",
                addinClass.AssemblyName, addinClass.AssemblyVersion, addinClass.AssemblyCulture, addinClass.AssemblyToken));
            content.AppendLine(String.Format("\"RuntimeVersion\"=\"{0}\"", addinClass.RuntimeVersion));

            if (Settings.Codebase)
                content.AppendLine(String.Format("\"Codebase\"=\"{0}\"", addinClass.Codebase));

            content.AppendLine(String.Empty);

            content.AppendLine("[" + addinClass.ClassesRoot + "\\CLSID\\{" + addinClass.Id + "}\\InprocServer32\\" + addinClass.AssemblyVersion + "]");
            content.AppendLine(String.Format("\"Class\"=\"{0}\"", addinClass.FullClassName));
            content.AppendLine(String.Format("\"Assembly\"=\"{0}, Version={1}, Culture={2}, PublicKeyToken={3}\"",
                addinClass.AssemblyName, addinClass.AssemblyVersion, addinClass.AssemblyCulture, addinClass.AssemblyToken));
            content.AppendLine(String.Format("\"RuntimeVersion\"=\"{0}\"", addinClass.RuntimeVersion));
            if (Settings.Codebase)
                content.AppendLine(String.Format("\"Codebase\"=\"{0}\"", addinAssembly.CodeBase));
            content.AppendLine(String.Empty);

            content.AppendLine("[" + addinClass.ClassesRoot + "\\CLSID\\{" + addinClass.Id + "}\\ProgId]");
            content.AppendLine(String.Format("@=\"{0}\"", addinClass.ProgId));
            content.AppendLine(String.Empty);

            content.AppendLine("[" + addinClass.ClassesRoot + "\\CLSID\\{" + addinClass.Id + "}\\Implemented Categories\\{" + addinClass.ComponentCategoryId.ToString() + "}]");

            if (Settings.AddinRegMode == SingletonSettings.AddinRegMode.On)
            {                
                Dictionary<object, Type> attributeTypes = GetAttributeTypes(addinClassType.GetCustomAttributes(true));
                bool isMultiAddin = AddinRegAnalyzer.IsMultiAddin(addinClassType);
                KeyValuePair<object, Type> multi = AddinRegAnalyzer.GetMultiRegisterAttribute(attributeTypes);
                KeyValuePair<object, Type> comAddin = AddinRegAnalyzer.GetComAddinAttribute(attributeTypes);
                if (null != comAddin.Key)
                {
                    string name = (string)comAddin.Value.InvokeMember("Name", BindingFlags.Instance | BindingFlags.Public | BindingFlags.GetField, null, comAddin.Key, new object[0]);
                    string description = (string)comAddin.Value.InvokeMember("Description", BindingFlags.Instance | BindingFlags.Public | BindingFlags.GetField, null, comAddin.Key, new object[0]);
                    int loadBehavior = (int)comAddin.Value.InvokeMember("LoadBehavior", BindingFlags.Instance | BindingFlags.Public | BindingFlags.GetField, null, comAddin.Key, new object[0]);
                    int commandLineSafe = (int)comAddin.Value.InvokeMember("CommandLineSafe", BindingFlags.Instance | BindingFlags.Public | BindingFlags.GetField, null, comAddin.Key, new object[0]);

                    if (true == isMultiAddin && multi.Key != null)
                    {
                        IEnumerable products = multi.Value.InvokeMember("Products", BindingFlags.Instance | BindingFlags.Public | BindingFlags.GetField, null, multi.Key, new string[0]) as IEnumerable;
                        if (null != products)
                        {
                            foreach (object item in products)
                            {
                                int productIndex = Convert.ToInt32(item);
                                CreateOfficeRegistryKey(content, _multiRegisterIn[productIndex], addinClass.ProgId,
                                    name, description, loadBehavior, commandLineSafe, Settings.Mode == SingletonSettings.RegisterMode.System);
                            }
                        }
                    }
                    else if (false == isMultiAddin)
                    {
                        string key = GetKeyName(addinClassType);
                        CreateOfficeRegistryKey(content, key, addinClass.ProgId, name, description, loadBehavior, commandLineSafe, Settings.Mode == SingletonSettings.RegisterMode.System);
                    }
                }
            }

            return content.ToString();
        }
        private static string _addinBaseClassName = "NetOffice.Tools.COMAddinBase";

        private static string _systemObject = "System.Object";

        private static string[] _classNames = new string[] { "NetOffice.MSProjectApi.Tools.COMAddin",
                                                        "NetOffice.ExcelApi.Tools.COMAddin",
                                                        "NetOffice.WordApi.Tools.COMAddin",
                                                        "NetOffice.OutlookApi.Tools.COMAddin",
                                                        "NetOffice.PowerPointApi.Tools.COMAddin",
                                                        "NetOffice.AccessApi.Tools.COMAddin",
                                                        "NetOffice.VisioApi.Tools.COMAddin",
                                                        "NetOffice.OfficeApi.Tools.COMAddin"};

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

        private static string[] _classKeys = new string[] { "MS Project", "Excel", "Word", "Outlook", "PowerPoint", "Access", "Visio" };

        private static string _officeRelatedKey = "Software\\Microsoft\\Office\\{0}\\Addins";

        private void CreateOfficeRegistryKey(StringBuilder builder, string officeKeyName, string addinProgId, string name, string description, int loadBehavior, int commandLineSafe, bool useSystemKey)
        {
            string hive = useSystemKey ? "HKEY_LOCAL_MACHINE" : "HKEY_CURRENT_USER";        
            string targetKey = String.Format(_officeRelatedKey + "\\{1}", officeKeyName, addinProgId);

            builder.AppendLine("");
            builder.AppendLine(String.Format(";--- {0} Addin ---", officeKeyName));
            builder.AppendLine(String.Format("[{0}\\{1}]", hive, targetKey));
            builder.AppendLine(String.Format("\"LoadBehavior\"=dword:{0}", RegValueConverter.ToDwordString(loadBehavior)));
            builder.AppendLine(String.Format("\"FriendlyName\"=\"{0}\"", name));
            builder.AppendLine(String.Format("\"Description\"=\"{0}\"", description));
            if(commandLineSafe > -1)
                builder.AppendLine(String.Format("\"CommandLineSafe\"=dword:{0}", RegValueConverter.ToDwordString(commandLineSafe)));    
        }

        private static Dictionary<object, Type> GetAttributeTypes(IEnumerable<object> attributes)
        {
            
            Dictionary<object, Type> result = new Dictionary<object, Type>();
            foreach (object item in attributes)
                result.Add(item, item.GetType());
            return result;
        }
        
        #region Events

        private Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            return Common.AssemblyResolve.Resolve(args.Name);
        }

        #endregion
    }
}
