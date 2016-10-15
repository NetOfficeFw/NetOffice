using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Security.Policy;
using System.Reflection;
using System.Runtime.InteropServices;
using RegAddin.Common;

namespace RegAddin.Register
{
    internal class RegisterOperationHandler
    {
        internal void Proceed()
        {
            AppDomainSetup domainSetup = new AppDomainSetup();
            domainSetup.ApplicationBase = Path.GetDirectoryName(SingletonSettings.AssemblyPath);
            domainSetup.LoaderOptimization = LoaderOptimization.MultiDomainHost;
            Evidence domainEvidence = AppDomain.CurrentDomain.Evidence;

            AppDomain domain = AppDomain.CreateDomain("RegisterDomain", domainEvidence, domainSetup);
            IAppDomainMethod appDomainInstance =
                (IAppDomainMethod)domain.CreateInstanceFromAndUnwrap(
                    typeof(IAppDomainMethod).Assembly.CodeBase, typeof(RegisterOperationHost).FullName);
                    
            appDomainInstance.SetConfig(
                new RegisterOperationHostSettings(
                    SingletonSettings.AssemblyPath, SingletonSettings.RegMode,
                    SingletonSettings.Codebase, SingletonSettings.DoRegisterCall == SingletonSettings.RegisterCall.On ? true : false,
                    SingletonSettings.SignCheck, SingletonSettings.Metrics, SingletonSettings.AddinReg));

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
    }
}
