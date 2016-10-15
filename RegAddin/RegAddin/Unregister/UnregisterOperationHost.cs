using System;
using System.Reflection;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.ComponentModel;
using System.Text;
using System.Runtime.InteropServices;
using RegAddin.Common;

namespace RegAddin.Unregister
{
    [Serializable]
    internal class UnregisterOperationHost : MarshalByRefObject, Common.IAppDomainMethod
    {
        #region Ctor

        public UnregisterOperationHost()
        {
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
        }

        #endregion

        #region Properties

        internal UnregisterOperationHostSettings Settings { get; private set; }

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
            UnregisterOperationHostSettings settings = configInstance as UnregisterOperationHostSettings;
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

            IEnumerable<object> assemblyAttributes = AssemblyReflection.GetCustomAssemblyAttributes(addinAssembly);
            Type[] types = addinAssembly.GetExportedTypes();
            foreach (Type item in types)
            {
                if (!item.IsClass)
                    continue;

                IEnumerable<object> addinClassAttributes = Common.AttributeReflection.GetCustomClassAttributes(item);
                if (!AddinClassReflection.IsValidAddinClass(addinClassAttributes, item.Attributes))
                    continue;

                AddinRegAnalyzer.KeyTarget target = AddinRegAnalyzer.KeyTarget.Both;
                switch (Settings.Mode)
                {
                    case SingletonSettings.UnRegisterMode.Auto:
                        target = AddinRegAnalyzer.KeyTarget.Both;
                        break;
                    case SingletonSettings.UnRegisterMode.System:
                        target = AddinRegAnalyzer.KeyTarget.System;
                        break;
                    case SingletonSettings.UnRegisterMode.User:
                        target = AddinRegAnalyzer.KeyTarget.User;
                        break;
                    default:
                        throw new IndexOutOfRangeException("Mode");
                }

                DeleteRegistryEntries(addinAssembly, assemblyAttributes, Settings.Mode, item, addinClassAttributes);
                DeleteOfficeKeys(item, addinClassAttributes, target);
                if (Settings.DoRegisterCall)
                {
                    int installScope = Settings.Mode == SingletonSettings.UnRegisterMode.System ? 0 : 1;
                    int keyState = Settings.AddinRegMode == SingletonSettings.AddinRegMode.Off ? 0 : 1;

                    if (!new Dispatcher.UnRegisterMethod().Call(item,
                        installScope,
                        keyState))
                        return (int)ResultCodes.UnRegisterCallFailed;
                }
            }

            return (int)ResultCodes.Okay;
        }

        private void DeleteOfficeKeys(Type addin, IEnumerable<object> addinClassAttributes, AddinRegAnalyzer.KeyTarget keyTarget)
        {
            if (Settings.AddinRegMode == SingletonSettings.AddinRegMode.On)
            {
                AddinRegAnalyzer reg = new AddinRegAnalyzer();
                reg.DeleteKey(addin, addinClassAttributes, keyTarget);
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

        private void DeleteRegistryEntries(Assembly addinAssembly, IEnumerable<object> assemblyAttributes, SingletonSettings.UnRegisterMode mode,
                       Type addinClassType, IEnumerable<object> addinClassAttributes)
        {
            AddinClassInformations addinClass = AddinClassInformations.Create(
                           addinAssembly, assemblyAttributes, mode, addinClassType, addinClassAttributes);

            Registry registry = new Registry();

            registry.DeleteComponentKey(Settings.Mode, addinClass.ProgId);
            registry.DeleteComponentKey(Settings.Mode, "CLSID\\{" + addinClass.Id + "}");            
        }

        #region Events

        private Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            return Common.AssemblyResolve.Resolve(args.Name);
        }

        #endregion
    }
}
