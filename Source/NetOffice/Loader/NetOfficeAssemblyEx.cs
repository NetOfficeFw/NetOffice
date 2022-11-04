using System;
using System.Collections.Generic;
using System.Reflection;
using NetOffice.Attributes;

namespace NetOffice.Loader
{
    /// <summary>
    /// Helper class to verify an assembly was signed with the NetOffice public key token.
    /// </summary>
    public static class NetOfficeAssemblyEx
    {
        /// <summary>
        /// Public key token of NetOffice libraries.
        /// </summary>
        public static readonly byte[] NetOfficePubliKeyToken = { 0x29, 0x7f, 0x57, 0xb4, 0x3a, 0xe7, 0xc1, 0xde };

        /// <summary>
        /// Returns true if the assembly has a NetOfficeAssemblyAttribute defined.
        /// </summary>
        /// <param name="itemAssembly">assembly</param>
        /// <returns>true if assembly has NetOfficeAssemblyAttribute, otherwise false</returns>
        internal static bool ContainsNetOfficeAttribute(this Assembly itemAssembly)
        {
            try
            {
                var attribute = itemAssembly.GetCustomAttribute<NetOfficeAssemblyAttribute>();
                return attribute != null;
            }
            catch (System.IO.FileNotFoundException)
            {
                return false;
            }
        }

        /// <summary>
        /// Returns true if the assembly is signed with NetOffice strong name.
        /// </summary>
        /// <param name="itemName">assembly information</param>
        /// <returns>true if NetOffice assembly with token, otherwise false</returns>
        internal static bool ContainsNetOfficePublicKeyToken(this AssemblyName itemName)
        {
            try
            {
                var token = itemName.GetPublicKeyToken();
                if (token == null || token.Length != NetOfficePubliKeyToken.Length)
                {
                    return false;
                }

                for (int i = 0; i < token.Length; i++)
                {
                    if (token[i] != NetOfficePubliKeyToken[i])
                    {
                        return false;
                    }
                }

                return true;
            }
            catch (System.IO.FileNotFoundException)
            {
                return false;
            }
        }
    }
}
