using System;
using System.IO;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.PowerPointApi.Tools.Contribution
{
    /// <summary>
    /// File related helper tools
    /// </summary>
    public class FileUtils
    {
        #region Fields

        private CommonUtils _owner;
        private static readonly string[] _extensions = new string[] { "pptx", "ppt", "pptm", "potx", "pot", "potm", "ppsx", "pps", "ppsm" };

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">owner instance</param>
        protected internal FileUtils(CommonUtils owner)
        {
            if (null == owner)
                throw new ArgumentNullException("owner");
            _owner = owner;
        }

        #endregion

        #region Methods

        /// <summary>
        /// Determines the given file name ends with a known powerpoint extension
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
                case "pptx":
                    return Contribution.FileExtension.Presentation;
                case "ppt":
                    return Contribution.FileExtension.PresentationDepricated;
                case "pptm":
                    return Contribution.FileExtension.PresentationInclMacros;
                case "potx":
                    return Contribution.FileExtension.Template;
                case "pot":
                    return Contribution.FileExtension.TemplateDepcricated;
                case "potm":
                    return Contribution.FileExtension.TemplateInclMacros;
                case "ppsx":
                    return Contribution.FileExtension.RuntimePresentation;
                case "pps":
                    return Contribution.FileExtension.RuntimePresentationDepricated;
                case "ppsm":
                    return Contribution.FileExtension.RuntimePresentationInclMacros;

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
                    return _owner.ApplicationIs2007OrHigher ? "pptx" : "ppt";
                case DocumentFormat.Macros:
                    return _owner.ApplicationIs2007OrHigher ? "pptm" : "ppt";
                case DocumentFormat.Template:
                    return _owner.ApplicationIs2007OrHigher ? "potx" : "pot";
                case DocumentFormat.TemplateMacros:
                    return _owner.ApplicationIs2007OrHigher ? "potm" : "pot";
                case DocumentFormat.Presentation:
                    return _owner.ApplicationIs2007OrHigher ? "ppsx" : "pps";
                case DocumentFormat.PresentationMacros:
                    return _owner.ApplicationIs2007OrHigher ? "ppsm" : "pps";
                default:
                    throw new ArgumentOutOfRangeException("type");
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