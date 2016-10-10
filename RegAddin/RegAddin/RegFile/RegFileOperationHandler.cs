using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Security.Policy;
using System.Reflection;
using System.Runtime.InteropServices;
using RegAddin.Common;

namespace RegAddin.RegFile
{  
    internal class RegFileOperationHandler
    {
        internal void Proceed()
        {                      
            AppDomainSetup domainSetup = new AppDomainSetup();
            domainSetup.ApplicationBase = Path.GetDirectoryName(SingletonSettings.AssemblyPath);
            domainSetup.LoaderOptimization = LoaderOptimization.MultiDomainHost;
                         Evidence domainEvidence = AppDomain.CurrentDomain.Evidence;

            AppDomain domain = AppDomain.CreateDomain("RegFileDomain", domainEvidence, domainSetup);
            IAppDomainMethod appDomainInstance =
                (IAppDomainMethod)domain.CreateInstanceFromAndUnwrap(
                    typeof(IAppDomainMethod).Assembly.CodeBase, typeof(RegFileOperationHost).FullName);
            
            appDomainInstance.SetConfig(
                new RegFileOperationHostSettings(
                    SingletonSettings.AssemblyPath, SingletonSettings.RegMode,
                    SingletonSettings.Codebase, SingletonSettings.RegFilePath, SingletonSettings.AddinReg, SingletonSettings.ExportCall));

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
