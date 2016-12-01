using System;
using System.IO;
using System.Reflection;
using System.Drawing;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Resource related utils
    /// </summary>
    public class ResourceUtils
    {
        #region Fields

        private CommonUtils _owner;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">owner instance</param>
        internal ResourceUtils(CommonUtils owner)
        {
            if (null == owner)
                throw new ArgumentNullException("owner");
            _owner = owner;
        }

        #endregion

        #region Methods

        /// <summary>
        /// Read stream from resource in executing assembly
        /// </summary>
        /// <param name="resourceAddress">full qualified resource address</param>
        /// <returns>resource stream</returns>
        public Stream ReadStream(string resourceAddress)
        {
            return ReadStream(resourceAddress, _owner.OwnerAssembly);
        }

        /// <summary>
        /// Read stream from resource in given assembly
        /// </summary>
        /// <param name="resourceAddress">full qualified resource address</param>
        /// <param name="assembly">target assembly to find resource</param>
        /// <returns>resource stream</returns>
        public Stream ReadStream(string resourceAddress, Assembly assembly)
        {
            System.IO.Stream resourceStream = assembly.GetManifestResourceStream(resourceAddress);
            if (resourceStream == null)
            {
                string target = _owner.Owner.GetType().Namespace + "." + resourceAddress;
                resourceStream = assembly.GetManifestResourceStream(target);
            }

            if (resourceStream == null)
                throw (new System.IO.IOException("Error accessing resource Stream."));

            return resourceStream;
        }

        /// <summary>
        /// Read string from resource in executing assembly
        /// </summary>
        /// <param name="resourceAddress">full qualified resource address</param>
        /// <returns>resource string</returns>
        public string ReadString(string resourceAddress)
        {
            return ReadString(resourceAddress, _owner.OwnerAssembly);
        }

        /// <summary>
        /// Read string from resource in given assembly
        /// </summary>
        /// <param name="resourceAddress">full qualified resource address</param>
        /// <param name="assembly">target assembly to find resource</param>
        /// <returns>resource string</returns>
        public string ReadString(string resourceAddress, Assembly assembly)
        {
            System.IO.Stream resourceStream = ReadStream(resourceAddress);
            System.IO.StreamReader textStreamReader = new System.IO.StreamReader(resourceStream);
            if (textStreamReader == null)
                throw (new System.IO.IOException("Error accessing resource string."));

            string text = textStreamReader.ReadToEnd();
            textStreamReader.Close();
            resourceStream.Close();
            return text;
        }

        /// <summary>
        /// Read image from resource in executing assembly
        /// </summary>
        /// <param name="resourceAddress">full qualified resource address</param>
        /// <returns>resource image</returns>
        public Image ReadImage(string resourceAddress)
        {
            Stream resourceStream = ReadStream(resourceAddress);
            return Bitmap.FromStream(resourceStream);
        }

        /// <summary>
        /// Read image from resource in given assembly
        /// </summary>
        /// <param name="resourceAddress">full qualified resource address</param>
        /// <param name="assembly">target assembly to find resource</param>
        /// <returns></returns>
        public Image ReadImage(string resourceAddress, Assembly assembly)
        {
            return ReadImage(resourceAddress, _owner.OwnerAssembly);
        }

        /// <summary>
        /// Read icon from resource in executing assembly
        /// </summary>
        /// <param name="resourceAddress">full qualified resource address</param>
        /// <returns>Resource icon</returns>
        public Icon ReadIcon(string resourceAddress)
        {
            return ReadIcon(resourceAddress, _owner.OwnerAssembly);
        }

        /// <summary>
        /// Read icon from resource in given assembly
        /// </summary>
        /// <param name="resourceAddress">full qualified resource address</param>
        /// <param name="assembly">target assembly to find resource</param>
        /// <returns>Resource icon</returns>
        public Icon ReadIcon(string resourceAddress, Assembly assembly)
        {
            Stream resourceStream = ReadStream(resourceAddress, assembly);            
            return new Icon(resourceStream);
        }

        #endregion
    }
}
