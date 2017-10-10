using System;
using Word = NetOffice.WordApi;
using NetOffice.OfficeApi.Enums;

namespace NetOffice.WordApi.Tools
{
    /// <summary>
    /// Provides DocumentInspector Services
    /// </summary>
    public abstract class ToolsDocumentInspector
    {
        /// <summary>
        /// Display Name
        /// </summary>
        public abstract string Name { get; }

        /// <summary>
        /// Display Description
        /// </summary>
        public abstract string Description { get; }

        /// <summary>
        /// Owner Addin
        /// </summary>
        protected COMAddin Owner { get; private set; }

        /// <summary>
        /// Initialize the instance 
        /// </summary>
        /// <param name="owner">owner addin</param>
        public virtual void InitializeInspector(COMAddin owner)
        {
            if (null == owner)
                throw new ArgumentNullException("owner");
            Owner = owner;
        }

        /// <summary>
        /// Gets information about a custom Document Inspector module.
        /// </summary>
        /// <param name="name">Represents the name of the module.</param>
        /// <param name="desc">Represents the description of the module.</param>
        public virtual void GetInfo(out string name, out string desc)
        {
            name = Name;
            desc = Description;
        }

        /// <summary>
        /// Inspects a document for specific information items or document properties by using a custom Document Inspector module.
        /// </summary>
        /// <param name="doc">An object representing the container document.</param>
        /// <param name="status">An MsoDocInspectorStatus value that represents the results of the inspection.</param>
        /// <param name="result">Contains a list of the information items or document properties found in the document.</param>
        /// <param name="action">Indicates to the user what action to take based on the results of the inspection.</param>
        public abstract void Inspect(Word.Document doc, out MsoDocInspectorStatus status,  out string result, out string action);

        /// <summary>
        /// Performs some action on specific information items or document properties by using a custom Document Inspector module.
        /// </summary>
        /// <param name="doc">Specifies an object representing the container object.</param>
        /// <param name="hwnd">Specifies the unique identifier of the active document window.</param>
        /// <param name="status">Specifies an enumeration that indicates the status of the action.</param>
        /// <param name="result">Contains the results of the action.</param>
        public abstract void Fix(Word.Document doc, int hwnd, out MsoDocInspectorStatus status, out string result);
    }
}