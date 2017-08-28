using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.CodeDom;
using Microsoft.CSharp;
using System.CodeDom.Compiler;
using NetOffice.Attributes;

namespace NetOffice.Duck
{
    internal class DuckTypeGenerator
    {
        private ulong _counter = ulong.MinValue;

        internal DuckTypeGenerator(DuckInterface proxyInterface)
        {
            if (null == proxyInterface)
                throw new ArgumentNullException("proxyInterface");
            ProxyInterface = proxyInterface;
        }

        private DuckInterface ProxyInterface { get; set; }

        internal Type GenerateType()
        {
            using (CodeDomProvider provider = new CSharpCodeProvider())
            {
                CompilerParameters args = new CompilerParameters();
                args.GenerateExecutable = false;
                args.GenerateInMemory = true;
                args.TreatWarningsAsErrors = false;
                args.IncludeDebugInformation = true;
                args.CompilerOptions = "/t:library /platform:anycpu /lib:" + "\"" +
                    System.IO.Path.GetDirectoryName(Resolver.UriResolver.ResolveLocalPath(typeof(ICOMObject).Assembly.Location))
                     + "\"";

                args.ReferencedAssemblies.Add("System.dll");
                args.ReferencedAssemblies.Add("mscorlib.dll");
                args.ReferencedAssemblies.Add("NetOffice.dll");
                args.ReferencedAssemblies.Add("IExcelApi.dll");
                
                args.ReferencedAssemblies.Add(ProxyInterface.AssemblyName);

                string[] modules = GenerateClassModule();
                CompilerResults result = provider.CompileAssemblyFromSource(args, modules);
                Type[] assemblyTypes = result.CompiledAssembly.GetTypes();
                return FindType(assemblyTypes);
            }
        }

        internal Type FindType(Type[] types)
        {
            foreach (Type item in types)
            {
                if (null != item.GetInterface(ProxyInterface.InterfaceType.FullName, false))
                    return item;
            }
            return null;
        }

        internal ulong AquireCounter()
        {
            if (ulong.MaxValue == _counter)
            {
                _counter = ulong.MinValue;
                return _counter;
            }
            else
            {
                _counter++;
                return _counter;
            }
        }

        internal string[] GenerateClassModule()
        {
            string name = (ProxyInterface.Name.StartsWith("I") ? ProxyInterface.Name.Substring(1) : ProxyInterface.Name) + "Duck_" + AquireCounter().ToString();

            StringBuilder classBuilder = new StringBuilder();
            StringBuilder issueBuilder = new StringBuilder();

            using (DuckTypeIssueClassGenerator issueClassGenerator = new DuckTypeIssueClassGenerator(issueBuilder, ProxyInterface, name))
            {
                using (DuckTypeClassGenerator classGenerator = new DuckTypeClassGenerator(classBuilder, ProxyInterface, name, issueClassGenerator.ImplementationName))
                {
                    if (ProxyInterface.IsValidEventClass)
                    {
                        EventInfo[] events = ProxyInterface.Events;
                        using (DuckTypeEventsGenerator eventsBuilder = new DuckTypeEventsGenerator(classBuilder, events))
                        {

                        }
                    }

                    HasIndexPropertyAttribute indexAttribute = ProxyInterface.GetHasIndexPropertyAttribute();
                    PropertyInfo[] indexProperties = ProxyInterface.PropertiesIndexer;
                    using (DuckTypeIndexerGenerator indexGenerator = new DuckTypeIndexerGenerator(classBuilder, indexProperties, indexAttribute))
                    {

                    }

                    PropertyInfo[] properties = ProxyInterface.Properties;
                    using (DuckTypePropertiesGenerator propertiesGenerator = new DuckTypePropertiesGenerator(classBuilder, properties))
                    {

                    }
                    
                 
                    MethodInfo[] methods = ProxyInterface.Methods;
                    using (DuckTypeMethodsGenerator methodsGenerator = new DuckTypeMethodsGenerator(classBuilder, methods))
                    {

                    }

                    EnumeratorAttribute enumAttribute = ProxyInterface.GetEnumeratorAttribute();
                    MethodInfo[] enumeratorMethods = ProxyInterface.MethodsWithEnumerator;
                    using (DuckTypeEnumeratorGenerator enumeratorGenerator = new DuckTypeEnumeratorGenerator(classBuilder, enumeratorMethods, enumAttribute))
                    {

                    }

                    MethodInfo[] issueMethods = ProxyInterface.MethodsWithSyntaxIssue;
                    using (DuckTypeMethodsGenerator methodsGenerator = new DuckTypeMethodsGenerator(issueBuilder, issueMethods))
                    {

                    }
                }
            }
          
            return new string[] { classBuilder.ToString(), issueBuilder.ToString() };
        }
    }
}
