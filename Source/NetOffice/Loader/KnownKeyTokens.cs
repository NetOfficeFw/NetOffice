using System;
using System.Collections.Generic;
using System.Reflection;

namespace NetOffice.Loader
{
    /// <summary>
    /// Contains embedded key token schema
    /// </summary>
    public class KnownKeyTokens : List<string>
    {
        /// <summary>
        /// Returns info the assembly is a NetOffice Api Assembly
        /// </summary>
        /// <param name="itemAssembly">assembly information</param>
        /// <returns>true if NetOffice assembly, otherwise false</returns>
        internal bool ContainsNetOfficeAttribute(Assembly itemAssembly)
        {
            try
            {
                List<string> dependAssemblies = new List<string>();
                object[] attributes = itemAssembly.GetCustomAttributes(true);
                foreach (object itemAttribute in attributes)
                {
                    string fullnameAttribute = itemAttribute.GetType().FullName;
                    if (fullnameAttribute == Attributes.NetOfficeAssemblyAttribute.FullName)
                        return true;
                }
                return false;
            }
            catch (System.IO.FileNotFoundException)
            {
                return false;
            }
        }

        /// <summary>
        /// Returns info the assembly is a NetOffice Api Assembly with known keytoken
        /// </summary>
        /// <param name="itemName">assembly information</param>
        /// <returns>true if NetOffice assembly with token, otherwise false</returns>
        internal bool ContainsNetOfficePublicKeyToken(AssemblyName itemName)
        {
            try
            {
                string targetKeyToken = itemName.FullName.Substring(itemName.FullName.LastIndexOf(" ") + 1);
                foreach (string item in this)
                {
                    if (item.EndsWith(targetKeyToken, StringComparison.InvariantCultureIgnoreCase))
                        return true;
                }
                return false;
            }
            catch (System.IO.FileNotFoundException)
            {
                return false;
            }
        }
    }
}
