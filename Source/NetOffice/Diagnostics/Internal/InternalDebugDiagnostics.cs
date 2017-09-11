using System;
using System.Reflection;
using System.Collections.Generic;
using System.Text;

#if DEBUG

namespace NetOffice.Diagnostics.Internal
{
    /// <summary>
    /// Internal selfdiagnostics/validation in debug build
    /// </summary>
    internal class InternalDebugDiagnostics
    {
        /// <summary>
        /// Validate attributation and key tokens 
        /// </summary>
        /// <param name="factory">core to check</param>
        internal void ValidateCore(Core factory)
        {
            ValidateKeyTokens(factory);
        }

        /// <summary>
        /// Validate key token versions match with core assembly version
        /// </summary>
        /// <param name="factory">core to check</param>
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
                        throw new NetOfficeException("Keytoken missmatch.");
                }
            }
        }

    }
}

#endif
