using System;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Runtime.Serialization.Formatters.Binary;
using System.Collections.Generic;
using System.CodeDom.Compiler;
using System.Text;

namespace NOTools.InMemoryCompiler
{
    /// <summary>
    /// Represents the definition for an In-Memory created assembly(.dll)
    /// </summary>
    public class DynamicAssembly
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="name">name of the assembly</param>
        /// <param name="references">references to other assemblies</param>
        public DynamicAssembly(string name, string[] references=null)
        {
            if (String.IsNullOrWhiteSpace(name))
                throw new ArgumentException("invalid assembly name");
            References = new DynamicReferencesCollection();
            Classes = new DynamicClassCollection(this);
            CustomClasses = new DynamicCustomClassCollection();
            Name = name;

            if (null != references)
                foreach (string item in references)
                    References.Add(item);
        }

        #endregion

        #region Properties

        /// <summary>
        /// Assembly References (System.Windows.Forms for example)
        /// </summary>
        public DynamicReferencesCollection References { get; private set; }

        /// <summary>
        /// Classes of the assembly
        /// </summary>
        public DynamicClassCollection Classes { get; private set; }

        /// <summary>
        /// Custom classes of the assembly
        /// </summary>
        public DynamicCustomClassCollection CustomClasses { get; private set; }

        /// <summary>
        /// Name of the assembly
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Additional path for Non-GAC reference assemblies
        /// </summary>
        public string ReferencesPath { get; set; }

        #endregion
    }
}
