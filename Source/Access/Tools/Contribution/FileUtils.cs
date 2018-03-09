using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using Access = NetOffice.AccessApi;

namespace NetOffice.AccessApi.Tools.Contribution
{
    /// <summary>
    /// File related helper tools
    /// </summary>
    public class FileUtils
    {
        #region Fields

        private bool _applicationIs2007OrHigher;
        private static readonly string[] _extensions = new string[] { "mdb", "accdb", "accde", "mde", "accdt", "mdt", "accda", "mda", "mdz", "accdf", "mdw" };

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">owner instance</param>
        /// <exception cref="ArgumentNullException">given owner is null</exception>
        protected internal FileUtils(CommonUtils owner)
        {
            if (null == owner)
                throw new ArgumentNullException("owner");
            _applicationIs2007OrHigher = owner.ApplicationIs2007OrHigher;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="applicationIs2007OrHigher">corresponding application version is 12 e.g. 2007 or higher</param>
        public FileUtils(bool applicationIs2007OrHigher)
        {
            _applicationIs2007OrHigher = applicationIs2007OrHigher;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="application"></param>
        /// <exception cref="ArgumentNullException">given application is null</exception>
        /// <exception cref="ObjectDisposedException">given application is already disposed</exception>
        public FileUtils(Access.Application application)
        {
            if (null == application)
                throw new ArgumentNullException("application");
            if (application.IsDisposed)
                throw new ObjectDisposedException("application");

            double? version = NetOffice.OfficeApi.Tools.Contribution.CommonUtils.TryGetApplicationVersion(application);
            if (null != version && version >= 12.00)
                _applicationIs2007OrHigher = true;
        }

        #endregion

        #region Methods

        /// <summary>
        /// Determines the given file name ends with a known access extension
        /// </summary>
        /// <param name="fileName">given file name</param>
        /// <returns>true if ends with extension, otherwise false</returns>
        public bool HasWellKnownExtension(string fileName)
        {
            if (String.IsNullOrWhiteSpace(fileName))
                return false;

            foreach (var item in _extensions)
            {
                if (fileName.EndsWith(item, StringComparison.InvariantCultureIgnoreCase))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Parse the extension for a given file name to determine the type. There is no warranties the file is truly what the extension supposed to be.
        /// </summary>
        /// <param name="fileName">given file name</param>
        /// <returns>file type or unknown</returns>
        /// <exception cref="ArgumentException">argument is null,empty orcontains invalid characters</exception>
        public FileExtension ParseExtension(string fileName)
        {
            if (String.IsNullOrWhiteSpace(fileName))
                throw new ArgumentException("Argument is null or empty.", "fileName");

            if (!ValidateNoInvalidCharacters(fileName))
                throw new ArgumentException("Argument contains one or more invalid characters.", "fileName");

            string extension = Path.GetExtension(fileName).ToLower().Trim();
            switch (extension)
            {
                case "accdb":
                    return Contribution.FileExtension.Database;
                case "mdb":
                    return Contribution.FileExtension.DatabaseDepricated;
                case "accde":
                    return Contribution.FileExtension.CompiledDatabase;
                case "mde":
                    return Contribution.FileExtension.CompiledDatabaseDepricated;
                case "accdr":
                    return Contribution.FileExtension.RuntimeDatabase;
                case "accdt":
                    return Contribution.FileExtension.Template;
                case "mdt":
                    return Contribution.FileExtension.TemplateDepricated;
                case "accda":
                    return Contribution.FileExtension.Addin;
                case "mda":
                    return Contribution.FileExtension.AddinDepcricated;
                case "mdz":
                    return Contribution.FileExtension.Assistant;
                case "accdf":
                    return Contribution.FileExtension.FieldDescription;
                case "mdw":
                    return Contribution.FileExtension.WorkgroupSecurity;
                default:
                    return Contribution.FileExtension.Unknown;
            }
        }

        /// <summary>
        /// Get the current default file extension for a document type. The method is not aware of the MS Compatibilty pack in 2003 or below
        /// </summary>
        /// <param name="type">target document type</param>
        /// <returns>default extension for document type</returns>
        public string FileExtension(DocumentFormat type)
        {
            switch (type)
            {
                case DocumentFormat.Normal:
                    return _applicationIs2007OrHigher ? "accdb" : "mdb";
                case DocumentFormat.Compiled:
                    return _applicationIs2007OrHigher ? "accde" : "mde";
                case DocumentFormat.Runtime:
                    return "accdr";
                case DocumentFormat.Template:
                    return _applicationIs2007OrHigher ? "accdt" : "mdt";
                case DocumentFormat.Addin:
                    return _applicationIs2007OrHigher ? "accda" : "mda";
                case DocumentFormat.Assistant:
                    return "mdz";
                case DocumentFormat.FieldDescription:
                    return "accdf";
                case DocumentFormat.WorkgroupSecurity:
                    return "mdw";
                default:
                    throw new ArgumentOutOfRangeException("<Unexpected document format. Please report this error.>", "type");
            }
        }

        /// <summary>
        /// Add dot extension to argument filename
        /// </summary>
        /// <param name="fileName">target file name</param>
        /// <param name="type">target document format</param>
        /// <returns>filename with dot and extension</returns>
        public string Combine(string fileName, DocumentFormat type)
        {
            string dotSeperator = fileName.EndsWith(".", StringComparison.InvariantCultureIgnoreCase) ? String.Empty : ".";
            return System.IO.Path.Combine(fileName, dotSeperator + FileExtension(type));
        }

        /// <summary>
        /// Combines 2 arguments and document type to valid file path
        /// </summary>
        /// <param name="directoryPath">target directory path</param>
        /// <param name="fileName">target file name</param>
        /// <param name="type">target document format</param>
        /// <returns>Combined file path</returns>
        public string Combine(string directoryPath, string fileName, DocumentFormat type)
        {
            string dotSeperator = fileName.EndsWith(".", StringComparison.InvariantCultureIgnoreCase) ? String.Empty : ".";
            return System.IO.Path.Combine(directoryPath, fileName + dotSeperator + FileExtension(type));
        }

        /// <summary>
        /// Checks arguments for invalid filesystem path characters
        /// </summary>
        /// <param name="value">given string as any</param>
        /// <returns>true if value is without invalid characters, otherwise false</returns>
        private bool ValidateNoInvalidCharacters(string value)
        {
            if (String.IsNullOrWhiteSpace(value))
                return true;
            char[] invalidChars = Path.GetInvalidPathChars();
            foreach (var item in invalidChars)
            {
                if (value.Contains(item.ToString()))
                    return false;
            }
            return true;
        }

        #endregion
    }
}