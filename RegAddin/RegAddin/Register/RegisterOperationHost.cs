using System;
using System.Reflection;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.ComponentModel;
using System.Text;
using System.Runtime.InteropServices;
using RegAddin.Common;

namespace RegAddin.Register
{
    [Serializable]
    internal class RegisterOperationHost : MarshalByRefObject, Common.IAppDomainMethod
    {
        #region Ctor

        public RegisterOperationHost()
        {
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
        }

        #endregion

        #region Properties

        internal RegisterOperationHostSettings Settings { get; private set; }

        private string AssemblyPath
        {
            get
            {
                return Path.GetDirectoryName(Settings.AssemblyPath);
            }
        }

        #endregion

        void Common.IAppDomainMethod.SetConfig(object configInstance)
        {
            RegisterOperationHostSettings settings = configInstance as RegisterOperationHostSettings;
            if (null != settings)
                Settings = settings;
            else
                throw new ArgumentException("Invalid configuration type.");
        }
  
        int Common.IAppDomainMethod.ExecuteInDomain()
        {
            AppDomain domain = AppDomain.CurrentDomain;
            Assembly addinAssembly = Assembly.LoadFile(Settings.AssemblyPath);
            int signCheckResult = DoTokenCheck(addinAssembly);
            if (0 != signCheckResult)
                return signCheckResult;

            DoMetrics(addinAssembly);

            IEnumerable<object> assemblyAttributes = AssemblyReflection.GetCustomAssemblyAttributes(addinAssembly);
            if (!AssemblyReflection.AssemblyIsComVisible(addinAssembly, assemblyAttributes))
                return (int)ResultCodes.AssemblyNotComVisible;

            Type[] types = addinAssembly.GetExportedTypes();
            foreach (Type item in types)
            {
                if (!item.IsClass)
                    continue;

                IEnumerable<object> addinClassAttributes = Common.AttributeReflection.GetCustomClassAttributes(item);
                if (!AddinClassReflection.IsValidAddinClass(addinClassAttributes, item.Attributes))
                    continue;

                CreateRegistryEntries(addinAssembly, assemblyAttributes, Settings.Mode, item, addinClassAttributes);
                CreateOfficeKeys(item, addinClassAttributes, Settings.Mode == SingletonSettings.RegisterMode.System);
                if (Settings.DoRegisterCall)
                {
                    if(!new Dispatcher.RegisterMethod().Call(item, 
                        Settings.Mode == SingletonSettings.RegisterMode.System ? 0 : 1,
                        Settings.AddinRegMode == SingletonSettings.AddinRegMode.Off ? 0 : 1))
                        return (int)ResultCodes.RegisterCallFailed;
                }
            }

            return (int)ResultCodes.Okay;
        }

        private void CreateOfficeKeys(Type addin, IEnumerable<object> addinClassAttributes, bool useSystemKey)
        {
            if (Settings.AddinRegMode == SingletonSettings.AddinRegMode.On)
            {
                AddinRegAnalyzer reg = new AddinRegAnalyzer();
                reg.CreateKey(addin, addinClassAttributes, useSystemKey);
            }
        }
        
        private void DoMetrics(Assembly addinAssembly)
        {
            if (Settings.Metrics != SingletonSettings.MetricsMode.None)
            {
                new Metrics.AddinMetrics(addinAssembly, Settings.Metrics == SingletonSettings.MetricsMode.Win).Check();
            }
        }

        private int DoTokenCheck(Assembly addinAssembly)
        {
            switch (Settings.SignCheck)
            {
                case SingletonSettings.SignCheckMode.Warn:
                    if (addinAssembly.GetName().GetPublicKeyToken().Length == 0)
                        new WarningPresenter().ShowWarning("The given assembly is not signed.");
                    break;
                case SingletonSettings.SignCheckMode.Error:
                    if (addinAssembly.GetName().GetPublicKeyToken().Length == 0)
                        return (int)ResultCodes.AssemblyNotSigned;
                    break;
                default:
                    break;
            }
            return 0;
        }

        private void CreateRegistryEntries(Assembly addinAssembly, IEnumerable<object> assemblyAttributes, SingletonSettings.RegisterMode mode,
                        Type addinClassType, IEnumerable<object> addinClassAttributes)
        {
            AddinClassInformations addinClass = AddinClassInformations.Create(
                           addinAssembly, assemblyAttributes, mode, addinClassType, addinClassAttributes);

            Registry registry = new Registry();

            Microsoft.Win32.RegistryKey key = registry.CreateComponentKey(Settings.Mode, addinClass.ProgId);
            registry.CreateComponentValue(key, "", addinClass.FullClassName, Microsoft.Win32.RegistryValueKind.String);
            key.Close();

            key =  registry.CreateComponentKey(Settings.Mode, String.Format("{0}\\CLSID", addinClass.ProgId));
            registry.CreateComponentValue(key, "", "{" + addinClass.Id.ToString() + "}", Microsoft.Win32.RegistryValueKind.String);
            key.Close();

            key = registry.CreateComponentKey(Settings.Mode, "CLSID\\{" + addinClass.Id + "}");
            registry.CreateComponentValue(key, "", addinClass.FullClassName, Microsoft.Win32.RegistryValueKind.String);
            key.Close();

            key = registry.CreateComponentKey(Settings.Mode, "CLSID\\{" + addinClass.Id + "}\\InprocServer32");
            registry.CreateComponentValue(key, "", "mscoree.dll", Microsoft.Win32.RegistryValueKind.String);
            registry.CreateComponentValue(key, "ThreadingModel", "Both", Microsoft.Win32.RegistryValueKind.String);
            registry.CreateComponentValue(key, "Class", addinClass.FullClassName, Microsoft.Win32.RegistryValueKind.String);
            registry.CreateComponentValue(key, "Assembly", String.Format("{0}, Version={1}, Culture={2}, PublicKeyToken={3}",
                addinClass.AssemblyName, addinClass.AssemblyVersion, addinClass.AssemblyCulture, addinClass.AssemblyToken), 
                Microsoft.Win32.RegistryValueKind.String);
            registry.CreateComponentValue(key, "RuntimeVersion", addinClass.RuntimeVersion, Microsoft.Win32.RegistryValueKind.String);
            if (Settings.Codebase)
                registry.CreateComponentValue(key, "Codebase", addinClass.Codebase, Microsoft.Win32.RegistryValueKind.String);
            key.Close();

            key = registry.CreateComponentKey(Settings.Mode, "CLSID\\{" + addinClass.Id + "}\\InprocServer32\\" + addinClass.AssemblyVersion);
            registry.CreateComponentValue(key, "Class", addinClass.FullClassName, Microsoft.Win32.RegistryValueKind.String);
            registry.CreateComponentValue(key, "Assembly", String.Format("{0}, Version={1}, Culture={2}, PublicKeyToken={3}",
               addinClass.AssemblyName, addinClass.AssemblyVersion, addinClass.AssemblyCulture, addinClass.AssemblyToken),
               Microsoft.Win32.RegistryValueKind.String);
            registry.CreateComponentValue(key, "RuntimeVersion", addinClass.RuntimeVersion, Microsoft.Win32.RegistryValueKind.String);
            if (Settings.Codebase)
                registry.CreateComponentValue(key, "Codebase", addinClass.Codebase, Microsoft.Win32.RegistryValueKind.String);
            key.Close();

            key = registry.CreateComponentKey(Settings.Mode, String.Format("CLSID\\{0}\\{1}","{" + addinClass.Id + "}", "ProgId"));
            registry.CreateComponentValue(key, "", addinClass.ProgId, Microsoft.Win32.RegistryValueKind.String);
            key.Close();

            key = registry.CreateComponentKey(Settings.Mode, String.Format("CLSID\\{0}\\Implemented Categories\\{1}", "{" + addinClass.Id + "}", "{" + addinClass.ComponentCategoryId +"}"));
            key.Close();            
        }

        #region Events

        private Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            return Common.AssemblyResolve.Resolve(args.Name);
        }

        #endregion
    }
}
