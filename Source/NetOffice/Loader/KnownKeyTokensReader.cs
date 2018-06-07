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
        /// Name the file we embedd as static resource
        /// </summary>
        private static string _keyTokensFileName = "KeyTokens.txt";

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

            using (System.IO.Stream ressourceStream = assembly.GetManifestResourceStream(coreType.Namespace + "." + _keyTokensFileName))
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
