using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using NetOffice;
using Office = NetOffice.OfficeApi;
using Word = NetOffice.WordApi;
using NetOffice.OfficeApi.Enums;
using NetOffice.Tools;
using NetOffice.OfficeApi.Tools;
using System.Reflection;
using NetOffice.Attributes;

namespace NetOffice.WordApi.Tools
{
    /// <summary>
    /// Encapsulate independent MS-Word IDocumentInspector.
    /// Need to have ProgId/Guid and DocumentInspector attribute in derived class.
    /// </summary>
    [ComVisible(true), ClassInterface(ClassInterfaceType.AutoDual)]
    public abstract class DocumentInspectorBase : Office.Native.IDocumentInspector
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public DocumentInspectorBase()
        {
            Core factory = CreateFactory();
            Factory = null != factory ? factory : Core.Default;
        }

        #endregion

        #region DocumentInspectorBase 

        /// <summary>
        /// Display Name
        /// </summary>
        protected abstract string Name { get; }

        /// <summary>
        /// Display Description
        /// </summary>
        protected abstract string Description { get; }
       
        /// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861133.aspx </remarks>
		/// <param name="doc">object doc</param>
		/// <param name="status">NetOffice.OfficeApi.Enums.MsoDocInspectorStatus status</param>
		/// <param name="result">string result</param>
		/// <param name="action">string action</param>
        protected abstract void Inspect(Word.Document doc, out MsoDocInspectorStatus status, out string result, out string action);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864114.aspx </remarks>
        /// <param name="doc">object doc</param>
        /// <param name="hwnd">Int32 hwnd</param>
        /// <param name="status">NetOffice.OfficeApi.Enums.MsoDocInspectorStatus status</param>
        /// <param name="result">string result</param>
        protected abstract void Fix(Word.Document doc, Int32 hwnd, out MsoDocInspectorStatus status, out string result);
       
        /// <summary>
        /// Factory Core
        /// </summary>
        protected Core Factory { get; private set; }

        /// <summary>
        /// Create the used factory. The method was called as first in the base ctor
        /// </summary>
        /// <returns>new Settings instance</returns>
        protected virtual Core CreateFactory()
        {
            Core core = new Core();
            return core;
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862465.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="desc">string desc</param>
        protected virtual void GetInfo(out string name, out string desc)
        {
            name = Name;
            desc = Description;
        }

        /// <summary>
        /// Try dispose given document after Inspect/Fix
        /// </summary>
        /// <param name="document">given document as any</param>
        /// <returns>true if no error occured, otherwise false</returns>
        protected virtual bool TryDisposeDocumentInspectorDocument(Word.Document document)
        {
            try
            {
                if (null != document && false == document.IsDisposed)
                {
                    document.Dispose();
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Unexpected error occured
        /// </summary>
        /// <param name="exception">error</param>
        /// <returns>true if error is marked as handled, otherwise false</returns>
        protected virtual bool OnBaseError(Exception exception)
        {
            return false;
        }

        #endregion

        #region IDocumentInspector

        void Office.Native.IDocumentInspector.GetInfo(out string name, out string desc)
        {
            try
            {
                GetInfo(out name, out desc);
            }
            catch (Exception exception)
            {
                name = null;
                desc = null;
                if (!OnBaseError(exception))
                    throw;
            }            
        }

        void Office.Native.IDocumentInspector.Inspect(object Doc, out MsoDocInspectorStatus Status, out string Result, out string Action)
        {
            try
            {
                Word.Document document = new Word.Document(Factory, null, Doc);
                try
                {
                    Inspect(document, out Status, out Result, out Action);
                }
                catch
                {
                    Status = MsoDocInspectorStatus.msoDocInspectorStatusError;
                    Result = null;
                    Action = null;
                    throw;
                }
                finally
                {
                    TryDisposeDocumentInspectorDocument(document);
                }
            }
            catch (Exception exception)
            {
                Status = MsoDocInspectorStatus.msoDocInspectorStatusError;
                Result = null;
                Action = null;
                if (!OnBaseError(exception))
                    throw;
            }
        }

        void Office.Native.IDocumentInspector.Fix(object Doc, Int32 Hwnd, out MsoDocInspectorStatus Status, out string Result)
        {
            try
            {
                Word.Document document = new Word.Document(Factory, null, Doc);
                try
                {
                    Fix(document, Hwnd, out Status, out Result);
                }
                catch
                {
                    Status = MsoDocInspectorStatus.msoDocInspectorStatusError;
                    Result = null;
                    throw;
                }
                finally
                {
                    TryDisposeDocumentInspectorDocument(document);
                }
            }
            catch (Exception exception)
            {
                Status = MsoDocInspectorStatus.msoDocInspectorStatusError;
                Result = null;
                if (!OnBaseError(exception))
                    throw;
            }
        }
        
        #endregion

        #region Register/Unregister

        /// <summary>
        /// Called from regasm while register 
        /// </summary>
        /// <param name="type">Type information for the class</param>
        [ComRegisterFunction, Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public static void RegisterFunction(Type type)
        {
            if (null == type)
                throw new ArgumentNullException("type");
            if (null != type.GetCustomAttribute<DontRegisterAddinAttribute>())
                return;

            RegisterHandleCodebase(type, InstallScope.System);
            RegisterHandleProgrammable(type, InstallScope.System);
            RegisterHandleDocumentInspector(type);
        }

        /// <summary>
        /// Called from regasm while ungregister
        /// </summary>
        /// <param name="type">Type information for the class</param>
        [ComUnregisterFunction, Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public static void UnregisterFunction(Type type)
        {
            if (null == type)
                throw new ArgumentNullException("type");
            if (null != type.GetCustomAttribute<DontRegisterAddinAttribute>())
                return;

            UnregisterHandleDocumentInspector(type);
            UnregisterHandleProgrammable(type, InstallScope.System);
            UnregisterHandleCodebase(type, InstallScope.System);
        }
        
        private static void RegisterHandleDocumentInspector(Type type)
        {
            try
            {
                DocumentInspectorAttribute[] attributes = DocumentInspectorAttribute.GetAttributes(type);
                if (attributes.Length > 0)
                {
                    GuidAttribute typeid = AttributeReflector.GetGuidAttribute(type);
                    foreach (var attribute in attributes)
                    {
                        foreach (var version in attribute.ProcessedApplicationVersion)
                        {
                            DocumentInspectorAttribute.CreateKey("Word", attribute.Name, version, attribute.Selected, typeid.Value);
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                if (!RegisterErrorHandler.RaiseStaticErrorHandlerMethod(type, RegisterErrorMethodKind.Register, exception))
                    throw;
            }
        }

        private static void UnregisterHandleDocumentInspector(Type type)
        {
            try
            {
                DocumentInspectorAttribute[] attributes = DocumentInspectorAttribute.GetAttributes(type);
                if (attributes.Length > 0)
                {
                    RegistryLocationAttribute location = AttributeReflector.GetRegistryLocationAttribute(type);
                    GuidAttribute typeid = AttributeReflector.GetGuidAttribute(type);
                    foreach (var attribute in attributes)
                    {
                        foreach (var version in attribute.ProcessedApplicationVersion)
                        {
                            DocumentInspectorAttribute.TryDeleteKey("Word", attribute.Name, version);
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                if (!RegisterErrorHandler.RaiseStaticErrorHandlerMethod(type, RegisterErrorMethodKind.UnRegister, exception))
                    throw;
            }
        }

        private static void RegisterHandleProgrammable(Type type, InstallScope scope)
        {
            try
            {
                RegistryLocationAttribute location = AttributeReflector.GetRegistryLocationAttribute(type);
                bool isSystemComponent = location.IsMachineComponentTarget(scope);
                ProgrammableAttribute programmable = AttributeReflector.GetProgrammableAttribute(type);
                if (null != programmable)
                {
                    ProgrammableAttribute.CreateKeys(type.GUID, isSystemComponent);

                }
            }
            catch (Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                if (!RegisterErrorHandler.RaiseStaticErrorHandlerMethod(type, RegisterErrorMethodKind.Register, exception))
                    throw;
            }
        }

        private static void UnregisterHandleProgrammable(Type type, InstallScope scope)
        {
            try
            {
                ProgrammableAttribute programmable = AttributeReflector.GetProgrammableAttribute(type);
                if (null != programmable)
                {
                    RegistryLocationAttribute location = AttributeReflector.GetRegistryLocationAttribute(type);
                    bool isSystemComponent = location.IsMachineComponentTarget(scope);
                    ProgrammableAttribute.DeleteKeys(type.GUID, isSystemComponent, false);
                }
            }
            catch (Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                if (!RegisterErrorHandler.RaiseStaticErrorHandlerMethod(type, RegisterErrorMethodKind.UnRegister, exception))
                    throw;
            }
        }

        private static void RegisterHandleCodebase(Type type, InstallScope scope)
        {
            try
            {
                CodebaseAttribute codebase = AttributeReflector.GetCodebaseAttribute(type);
                if (null != codebase && codebase.Value)
                {
                    RegistryLocationAttribute location = AttributeReflector.GetRegistryLocationAttribute(type);
                    bool isSystemComponent = location.IsMachineComponentTarget(scope);
                    Assembly thisAssembly = Assembly.GetAssembly(type);
                    string assemblyVersion = thisAssembly.GetName().Version.ToString();
                    CodebaseAttribute.CreateValue(type.GUID, isSystemComponent, assemblyVersion, thisAssembly.CodeBase);

                }
            }
            catch (Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                if (!RegisterErrorHandler.RaiseStaticErrorHandlerMethod(type, RegisterErrorMethodKind.Register, exception))
                    throw;
            }
        }

        private static void UnregisterHandleCodebase(Type type, InstallScope scope)
        {
            try
            {
                CodebaseAttribute codebase = AttributeReflector.GetCodebaseAttribute(type);
                if (null != codebase && codebase.Value == true)
                {
                    RegistryLocationAttribute location = AttributeReflector.GetRegistryLocationAttribute(type);
                    bool isSystemComponent = location.IsMachineComponentTarget(scope);
                    Assembly thisAssembly = Assembly.GetAssembly(type);
                    string assemblyVersion = thisAssembly.GetName().Version.ToString();
                    CodebaseAttribute.DeleteValue(type.GUID, isSystemComponent, assemblyVersion, false);
                }
            }
            catch (Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                if (!RegisterErrorHandler.RaiseStaticErrorHandlerMethod(type, RegisterErrorMethodKind.UnRegister, exception))
                    throw;
            }
        }

        #endregion
    }
}