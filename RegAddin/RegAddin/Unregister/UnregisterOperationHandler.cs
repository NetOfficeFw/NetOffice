using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Security.Policy;
using System.Reflection;
using System.Runtime.InteropServices;
using RegAddin.Common;

namespace RegAddin.Unregister
{
    internal class UnregisterOperationHandler
    { 
        internal void Proceed()
        {
            if (!PreCheckIsActionRequired())
                return;

            AppDomainSetup domainSetup = new AppDomainSetup();
            domainSetup.ApplicationBase = Path.GetDirectoryName(SingletonSettings.AssemblyPath);
            domainSetup.LoaderOptimization = LoaderOptimization.MultiDomainHost;
            Evidence domainEvidence = AppDomain.CurrentDomain.Evidence;

            AppDomain domain = AppDomain.CreateDomain("UnregisterDomain", domainEvidence, domainSetup);
            IAppDomainMethod appDomainInstance =
                (IAppDomainMethod)domain.CreateInstanceFromAndUnwrap(
                    typeof(IAppDomainMethod).Assembly.CodeBase, typeof(UnregisterOperationHost).FullName);

            appDomainInstance.SetConfig(
                new UnregisterOperationHostSettings(
                    SingletonSettings.AssemblyPath, SingletonSettings.UnRegMode,
                    SingletonSettings.DoRegisterCall == SingletonSettings.RegisterCall.On ? true : false,
                    SingletonSettings.SignCheck, SingletonSettings.AddinReg));

            int result = 0;
            try
            {
                result = appDomainInstance.ExecuteInDomain();
            }
            catch (Exception)
            {
                throw;
            }

            new ResultValidator(result).ThrowIfNeeded();

            try
            {
                AppDomain.Unload(domain);
            }
            catch
            {
                ;
            }
        }

        private bool PreCheckIsActionRequired()
        {
            if (!File.Exists(SingletonSettings.AssemblyPath))
                return !SingletonSettings.SuspendMissingAssemblyErrorInUnregister;
            else
                return true;
        }
    }
}
