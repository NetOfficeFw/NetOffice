using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Native;
using NetOffice.Tools;

namespace NetOffice.OfficeApi.Tools
{
    /// <summary>
    /// RibbonExtensibility base to seperate ribbon logics from addin connect
    /// </summary>
    [ComVisible(true), ClassInterface(ClassInterfaceType.AutoDual)]
    public class RibbonExtensibility : NetOffice.OfficeApi.Native.IRibbonExtensibility
    {
        #region Fields

        private Type _type;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="parent">addin owner instance</param>
        /// <exception cref="ArgumentNullException">parent is null(Nothing in Visual Basic)</exception>
        public RibbonExtensibility(IOfficeCOMAddin parent)
        {
            if (null == parent)
                throw new ArgumentNullException("parent");
            Parent = parent;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Addin Owner Instance
        /// </summary>
        public IOfficeCOMAddin Parent { get; set; }

        /// <summary>
        /// Ribbon instance to manipulate ui at runtime
        /// </summary>
        public IRibbonUI RibbonUI { get; private set; }

        /// <summary>
        /// Used Factory Core
        /// </summary>
        protected internal Core Factory
        {
            get
            {
                return null != Parent ? Parent.Factory : Core.Default;
            }
        }

        /// <summary>
        /// Instance Type (cached)
        /// </summary>
        protected internal Type Type
        {
            get
            {
                if (null == _type)
                    _type = GetType();
                return _type;
            }
        }

        #endregion

        #region IRibbonExtensibility

        /// <summary>
        /// IRibbonExtensibility implementation
        /// </summary>
        /// <param name="RibbonID">target ribbon id</param>
        /// <returns>XML content or String.Empty</returns>
        string IRibbonExtensibility.GetCustomUI(string RibbonID)
        {
            try
            {
                return OnGetCustomUI(RibbonID);
            }
            catch (Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                OnError(ErrorMethodKind.GetCustomUI, exception);
                return string.Empty;
            }
        }

        /// <summary>
        /// IRibbonExtensibility implementation
        /// </summary>
        /// <param name="ribbonID">target ribbon id</param>
        /// <returns>XML content or String.Empty</returns>
        protected internal virtual string OnGetCustomUI(string ribbonID)
        {
            var ribbon = NetOffice.Attributes.AttributeExtensions.GetCustomAttribute<CustomUIAttribute>(Type);
            if (null != ribbon && ribbon.RibbonIDs.Contains(ribbonID))
                return ReadString(CustomUIAttribute.BuildPath(ribbon.Value, ribbon.UseAssemblyNamespace, Type.Namespace));
            else
                return string.Empty;
        }

        /// <summary>
        /// Pre-defined Ribbon Loader
        /// </summary>
        /// <param name="ribbonUI">actual ribbon ui</param>
        public virtual void CustomUI_OnLoad(Office.Native.IRibbonUI ribbonUI)
        {
            try
            {
                RibbonUI = COMObject.Create<OfficeApi.IRibbonUI>(Factory, ribbonUI);
            }
            catch (Exception exception)
            {
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.GetCustomUI, exception);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Read string from resource in executing assembly
        /// </summary>
        /// <param name="resourceAddress">full qualified resource address</param>
        /// <returns>resource string</returns>
        private string ReadString(string resourceAddress)
        {
            return ReadString(resourceAddress, Parent.Type.Assembly);
        }

        /// <summary>
        /// Read string from resource in given assembly
        /// </summary>
        /// <param name="resourceAddress">full qualified resource address</param>
        /// <param name="assembly">target assembly to find resource</param>
        /// <returns>resource string</returns>
        private string ReadString(string resourceAddress, Assembly assembly)
        {
            System.IO.Stream resourceStream = ReadStream(resourceAddress, assembly);
            System.IO.StreamReader textStreamReader = new System.IO.StreamReader(resourceStream);
            if (textStreamReader == null)
                throw (new System.IO.IOException("Error accessing resource string."));

            string text = textStreamReader.ReadToEnd();
            textStreamReader.Close();
            resourceStream.Close();
            return text;
        }

        /// <summary>
        /// Read stream from resource in given assembly
        /// </summary>
        /// <param name="resourceAddress">full qualified resource address</param>
        /// <param name="assembly">target assembly to find resource</param>
        /// <returns>resource stream</returns>
        private Stream ReadStream(string resourceAddress, Assembly assembly)
        {
            System.IO.Stream resourceStream = assembly.GetManifestResourceStream(resourceAddress);
            if (resourceStream == null)
            {
                string target = Parent.Type.Namespace + "." + resourceAddress;
                resourceStream = assembly.GetManifestResourceStream(target);
            }

            if (resourceStream == null)
                throw (new System.IO.IOException("Error accessing resource Stream."));

            return resourceStream;
        }

        #endregion

        #region Trigger

        /// <summary>
        /// Custom error handler
        /// </summary>
        /// <param name="methodKind">origin method where the error comes from</param>
        /// <param name="exception">occured exception</param>
        protected virtual void OnError(ErrorMethodKind methodKind, Exception exception)
        {

        }

        #endregion
    }
}
