using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Loader
{
    /// <summary>
    /// Read known tokes from embedded resource
    /// </summary>
    internal class KnownKeyTokensReader
    {
        /// <summary>
        /// Name the file we embedd as static resource in Debug Build
        /// </summary>
        private string _keyTokensFileNameDebug = "KeyTokens_Debug.txt";

        /// <summary>
        /// Name the file we embedd as static resource in Release Build
        /// </summary>
        private string _keyTokensFileNameRelease = "KeyTokens_Release.txt";

        /// <summary>
        /// Perform reading
        /// </summary>
        /// <returns>key token sequence</returns>
        internal KnownKeyTokens Read()
        {
            string[] tokens = KeyTokens();
            KnownKeyTokens knownNetOfficeKeyTokens = new KnownKeyTokens();
            foreach (string item in tokens)
                knownNetOfficeKeyTokens.Add(item);
            return knownNetOfficeKeyTokens;
        }

        /// <summary>
        /// Returns embedded keytoken schema
        /// </summary>
        /// <returns>keytoken line array</returns>
        internal string[] KeyTokens()
        {
            var coreType = typeof(Core);
            var assembly = coreType.Assembly;

            #if DEBUG
                string keyTokensFile = _keyTokensFileNameDebug;
            #else
                string keyTokensFile = _keyTokensFileNameRelease;
            #endif

            using (System.IO.Stream ressourceStream = assembly.GetManifestResourceStream(coreType.Namespace + "." + keyTokensFile))
            {
                using (System.IO.StreamReader textStreamReader = new System.IO.StreamReader(ressourceStream))
                {
                    string text = textStreamReader.ReadToEnd();
                    return text.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                }
            }
        }
    }
}
