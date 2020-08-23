using System;
using System.IO;
using System.Reflection;
using System.Drawing;

namespace NetOffice.OfficeApi.Tools.Contribution
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
        /// Read bytes from resource in executing assembly
        /// </summary>
        /// <param name="resourceAddress">full qualified resource address</param>
        /// <returns>resource stream</returns>
        public byte[] ReadBytes(string resourceAddress)
        {
            using (Stream stream = ReadStream(resourceAddress, _owner.OwnerAssembly))
            {
                byte[] bytes = new byte[stream.Length];
                stream.Read(bytes, 0, Convert.ToInt32(stream.Length));
                return bytes;
            }
        }

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
            {
                throw new IOException($"Resource '{resourceAddress}' does not exists in assembly '{assembly.GetName().Name}.");
            }

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
            using (var resourceStream = ReadStream(resourceAddress, assembly))
            using (var textStreamReader = new StreamReader(resourceStream))
            {
                var text = textStreamReader.ReadToEnd();
                return text;
            }
        }

        /// <summary>
        /// Read image from resource in executing assembly
        /// </summary>
        /// <param name="resourceAddress">full qualified resource address</param>
        /// <returns>resource image</returns>
        public Image ReadImage(string resourceAddress)
        {
            return ReadImage(resourceAddress, _owner.OwnerAssembly);
        }

        /// <summary>
        /// Read image from resource in given assembly
        /// </summary>
        /// <param name="resourceAddress">full qualified resource address</param>
        /// <param name="assembly">target assembly to find resource</param>
        /// <returns></returns>
        public Image ReadImage(string resourceAddress, Assembly assembly)
        {
            Stream resourceStream = ReadStream(resourceAddress, assembly);
            return Image.FromStream(resourceStream);
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
