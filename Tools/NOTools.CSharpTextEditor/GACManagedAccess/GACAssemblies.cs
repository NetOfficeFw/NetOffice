using System;
using System.Threading;
using System.Reflection;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.CSharpTextEditor.GACManagedAccess
{
    internal delegate void LoadAssemblyUpdateEventHandler(GACAssembly[] gacAssemblies, bool isLastUpdate);

    internal class GACAssemblies  : List<GACAssembly>
    {
        private static void RunAsyncMethod(ThreadStart method)
        {
            Thread thread1 = new Thread(method);
            thread1.Start();
        }

        public static void BeginLoadAssemblyInformations(LoadAssemblyUpdateEventHandler updateHandler, int resultSize, string netRuntimeCondition = "NET_4_0")
        {
            RunAsyncMethod(
               delegate
               {
                   GACAssemblies list = new GACAssemblies();

                   AssemblyCacheEnum cache = new AssemblyCacheEnum(null);
                   string assemblyString = cache.GetNextAssembly();

                   string assemblyName = string.Empty;
                   if (!String.IsNullOrEmpty(assemblyString))
                   {
                       GACAssembly gacAssembly = QueryAssemblyInfo(assemblyString);
                       if (CheckConditions(gacAssembly, netRuntimeCondition))
                           list.Add(gacAssembly);
                   }
                   else
                   {
                       updateHandler(list.ToArray(), true);
                       return;
                   }

                   while (null != assemblyString)
                   {
                       assemblyString = cache.GetNextAssembly();
                       if (!String.IsNullOrEmpty(assemblyString))
                       {
                           GACAssembly gacAssembly = QueryAssemblyInfo(assemblyString);
                           if (!CheckConditions(gacAssembly, netRuntimeCondition))
                               continue;
                           list.Add(gacAssembly);
                       }

                       if (list.Count >= resultSize)
                       {
                           updateHandler(list.ToArray(), false);
                           list.Clear();
                       }
                   }

                   updateHandler(list.ToArray(), true);
                   list.Clear();
               });
        }

        public static GACAssemblies LoadAssemblyInformations(bool sortAlphabetically = true, string netRuntimeCondition = "NET_4_0")
        {
            GACAssemblies list = new GACAssemblies();

            AssemblyCacheEnum cache = new AssemblyCacheEnum(null);
            string assemblyString = cache.GetNextAssembly();
        
            string assemblyName = string.Empty;
            if (!String.IsNullOrEmpty(assemblyString))
            {
                GACAssembly gacAssembly = QueryAssemblyInfo(assemblyString);
                if (CheckConditions(gacAssembly, netRuntimeCondition))
                    list.Add(gacAssembly);
            }
            else
                return list;

            while (null != assemblyString)
            {
                assemblyString = cache.GetNextAssembly();
                if (!String.IsNullOrEmpty(assemblyString))
                {
                    GACAssembly gacAssembly = QueryAssemblyInfo(assemblyString);
                    if (!CheckConditions(gacAssembly, netRuntimeCondition))
                        continue;
                    list.Add(gacAssembly);     
                }
            }

            if (sortAlphabetically)
                list.Sort(new GACAssemblyComparer());

            return list;
        }

        private static GACAssembly QueryAssemblyInfo(string assemblyString)
        {
            string assemblyPath = AssemblyCache.QueryAssemblyInfo(assemblyString);
            Mono.Cecil.AssemblyDefinition assemblyDefinition = Mono.Cecil.AssemblyDefinition.ReadAssembly(assemblyPath);
            return new GACAssembly(assemblyDefinition.Name.Name,
                                    assemblyDefinition.Name.Version,
                                    assemblyPath, 
                                    assemblyDefinition.MainModule.Runtime.ToString(), 
                                    KeyTokenToString(assemblyDefinition.Name.PublicKeyToken));
        }

        private static bool CheckConditions(GACAssembly gacAssembly, string netRuntimeCondition = "NET_4_0")
        {
            if (String.IsNullOrWhiteSpace(netRuntimeCondition))
                return true;
            return (String.Equals(gacAssembly.Runtime ,netRuntimeCondition, StringComparison.InvariantCultureIgnoreCase));
        }

        private static string KeyTokenToString(byte[] pt)
        {
            string result = String.Empty;
            for (int i = 0; i < pt.GetLength(0); i++)
                result += String.Format("{0:x2}", pt[i]);
            return result;
        }
    }
}
