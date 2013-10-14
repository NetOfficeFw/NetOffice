using System;
using System.Reflection;
using System.Collections.Generic;
using System.CodeDom.Compiler;
using System.Text;

namespace NOTools.InMemoryCompiler
{
    /// <summary>
    /// Complex return value for CSharpCompiler.cs => CompileDynamicAssembly
    /// </summary>
    public class CompileResult
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="codeModules">source Code modules(classes)</param>
        /// <param name="errors">compiler errors</param>
        /// <param name="assembly">compiled assemby(if sucseed)</param>
        internal CompileResult(string[] codeModules, CompilerErrorCollection errors, Assembly assembly)
        {
            CodeModules = codeModules;
            Errors = errors;
            Assembly = assembly;
        }

        /// <summary>
        /// Compiled Assembly (if sucseed)
        /// </summary>
        public Assembly Assembly { get; private set; }

        /// <summary>
        /// Source Code modules(classes)
        /// </summary>
        public string[] CodeModules { get; private set; }

        /// <summary>
        /// Compiler Errors
        /// </summary>
        public CompilerErrorCollection Errors { get; private set; }
    }
}
