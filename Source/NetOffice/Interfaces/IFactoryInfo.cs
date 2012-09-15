using System;
using System.Reflection; 
using System.ComponentModel;
using System.Collections.Generic;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// info about a NetOffice assembly
    /// </summary>
    public interface IFactoryInfo
    {
        /// <summary>
        /// namespace of assembly
        /// </summary>
        string AssemblyNamespace { get; }

        /// <summary>
        /// guid of component there represents the NetOfficeApi assembly
        /// </summary>
        Guid ComponentGuid { get; }
        
        /// <summary>
        /// assembly info of NetOfficeApi assembly
        /// </summary>
        Assembly Assembly { get; }

        /// <summary>
        /// returns info a class with given name exists in NetOfficeApi assembly
        /// </summary>
        /// <param name="className"></param>
        /// <returns></returns>
        bool Contains(string className);

        /// <summary>
        /// returns a name array of dependent NetOfficeApi assemblies
        /// </summary>
        string[] Dependencies { get; }
    }
}
