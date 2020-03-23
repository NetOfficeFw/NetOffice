using System;
using System.Reflection;
using System.Collections.Generic;
using System.Text;

#if DEBUG

namespace NetOffice.Diagnostics.Internal
{
    /// <summary>
    /// Internal self diagnostics and validation in debug builds.
    /// </summary>
    internal class InternalDebugDiagnostics
    {
        /// <summary>
        /// Validates attributes and key tokens.
        /// </summary>
        /// <param name="factory">core to check</param>
        internal void ValidateCore(Core factory)
        {
            ValidateKeyTokens(factory);
        }

        /// <summary>
        /// Validates key token versions matches with the version of the Core assembly.
        /// </summary>
        /// <param name="factory">Core factory object to check</param>
        private void ValidateKeyTokens(Core factory)
        {
            string version = "Version=" + factory.ThisAssembly.GetName().Version.ToString();
            string[] keyTokens = NetOffice.Loader.CurrentAppDomain.KeyTokens(factory);
            string[] splitter = new string[] { "," };
            foreach (string line in keyTokens)
            {
                if (line.StartsWith(";"))
                    continue;
                string[] lineArray = line.Split(splitter, StringSplitOptions.RemoveEmptyEntries);
                if (lineArray.Length == 4)
                {
                    string lineVersion = lineArray[1].Trim();
                    if (version != lineVersion)
                        throw new NetOfficeException("Key tokens from KeyTokens.txt file does not match the version of NetOffice Core assembly.");
                }
            }
        }

    }
}

#endif
