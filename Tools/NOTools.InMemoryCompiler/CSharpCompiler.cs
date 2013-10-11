using System;
using System.Collections.Generic;
using System.CodeDom.Compiler;
using System.Text;

namespace NOTools.InMemoryCompiler
{
    /// <summary>
    /// C# Compiler class to create In-Memory .dll assemblies
    /// </summary>
    public static class CSharpCompiler
    {
        /// <summary>
        /// Creates an Assembly from DynamicAssembly definition
        /// </summary>
        /// <param name="assemblyDefinition">assembly description</param>
        /// <returns>result with assembly or error info</returns>
        public static CompileResult CompileDynamicAssembly(DynamicAssembly assemblyDefinition)
        {
            CodeDomProvider codeDomProvider = CodeDomProvider.CreateProvider("CSharp");
            CompilerParameters compilerParameters = new CompilerParameters();
            compilerParameters.CompilerOptions = "/t:library /platform:anycpu /lib:" + "\"" + 
                (String.IsNullOrWhiteSpace(assemblyDefinition.ReferencesPath) ? GetCurrentPath() : assemblyDefinition.ReferencesPath) + "\"";
            compilerParameters.IncludeDebugInformation = false;
            //compilerParameters.OutputAssembly = assemblyDefinition.Name;
            compilerParameters.GenerateInMemory = true;
            compilerParameters.GenerateExecutable = false;   
            
            foreach (var item in assemblyDefinition.References)
                compilerParameters.ReferencedAssemblies.Add(item);

            List<string> codeModules = new List<string>();
            foreach (DynamicClass item in assemblyDefinition.Classes)
            {
                string code = DynamicClass.Template;
                code = DynamicClass.AddUsingsToTemplate(item, code);
                code = DynamicClass.AddInterfacesToTemplate(item, code);
                code = DynamicClass.AddNamespaceToTemplate(String.IsNullOrWhiteSpace(item.Namespace) ? assemblyDefinition.Name : item.Namespace, code);
                code = DynamicClass.AddNameToTemplate(item, code);
                code = DynamicClass.AddPropertiesToTemplate(item, code);
                code = DynamicClass.AddMethodsToTemplate(item, code);
                codeModules.Add(code);
            }

            foreach (DynamicCustomClass item in assemblyDefinition.CustomClasses)
                codeModules.Add(item.Code);

            // we dont allow empty class definitions(fun fact: its okay for the c# compiler)
            foreach (DynamicCustomClass item in assemblyDefinition.CustomClasses)
            {
                if (String.IsNullOrWhiteSpace(item.Code))
                {
                    CompilerErrorCollection collection = new CompilerErrorCollection();
                    CompilerError customError = new CompilerError("CustomClass", 0, 0, "Custom", "Unable to compile an empty code module.");
                    collection.Add(customError);
                    return new CompileResult(codeModules.ToArray(), collection, null);
                }
            }

            CompilerResults compilerResults = codeDomProvider.CompileAssemblyFromSource(compilerParameters, codeModules.ToArray());
            codeDomProvider.Dispose();

            return new CompileResult(codeModules.ToArray(), compilerResults.Errors, compilerResults.Errors.Count > 0 ? null : compilerResults.CompiledAssembly);
        }

        /// <summary>
        /// Returns current codebase folder
        /// </summary>
        /// <returns>Folder path</returns>
        private static string GetCurrentPath()
        {
            string fileName = typeof(CSharpCompiler).Assembly.Location;
            return System.IO.Path.GetDirectoryName(fileName);
        }
    }
}
