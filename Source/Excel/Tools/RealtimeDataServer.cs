using System;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using NetOffice.Attributes;
using NetOffice.Tools;
using NetOffice.Exceptions;
using NetOffice.ExcelApi.Tools.Attributes;
using NetOffice.ExcelApi.Tools.Exceptions;

namespace NetOffice.ExcelApi.Tools
{
    /// <summary>
    /// NetOffice Realtime Data Server(IRtdServer) in MS-Excel
    /// SupportByVersion Excel, 10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
    public abstract class RealtimeDataServer : Native.IRtdServer
    {
        #region Fields

        private object _thisLock = new object();
        private COMRtdServerAttribute _attribute;
        private Type _instanceType;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public RealtimeDataServer()
        {
            try
            {
                Initialize();
            }
            catch (Exception exception)
            {
                OnError(RTDMethods.Unknown, exception);
                throw new COMRtdServerException(exception);
            }
        }

        #endregion

        #region Properties

        /// <summary>
        /// Used factory core
        /// </summary>
        protected internal Core Factory { get; private set; }

        /// <summary>
        /// Cached instance type
        /// </summary>
        protected internal Type InstanceType
        {
            get
            {
                if (null == _instanceType)
                    _instanceType = GetType();
                return _instanceType;
            }
        }

        /// <summary>
        /// Cached RtdServer attribute
        /// </summary>
        protected internal COMRtdServerAttribute Attribute
        {
            get
            {
                if (null == _attribute)
                    _attribute = InstanceType.GetCustomAttribute<COMRtdServerAttribute>();
                return _attribute;
            }
        }

        /// <summary>
        /// Notify callback given in ServerStart
        /// </summary>
        protected IRTDUpdateEvent CallbackObject { get; set; }

        #endregion

        #region IRtdServer Virtuals

        /// <summary>
        /// SupportByVersion Excel, 10,11,12,14,15,16
        /// The ServerStart method is called immediately after a real-time data server is instantiated. Negative value or zero indicates failure to start the server; positive value indicates success.
        /// </summary>
        /// <param name="callbackObject">IRTDUpdateEvent object. The callback object.</param>
        /// <returns>System.Int32</returns>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        protected virtual int ServerStart(IRTDUpdateEvent callbackObject)
        {
            if (null != callbackObject)
            {
                CallbackObject = callbackObject;
                return 1;
            }

            else
                return -1;
        }

        /// <summary>
        /// SupportByVersion Excel, 10,11,12,14,15,16
        /// Adds new topics from a real-time data server. The ConnectData method is called when a file is opened that contains real-time data functions or when a user types in a new formula which contains the RTD function.
        /// </summary>
        /// <param name="topicID">A unique value, assigned by Microsoft Excel, which identifies the topic</param>
        /// <param name="strings">A single-dimensional array of strings identifying the topic</param>
        /// <param name="getNewValues">True to determine if new values are to be acquired</param>
        /// <returns>System.Object</returns>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        protected abstract object ConnectData(int topicID, object strings, bool getNewValues);
        
        /// <summary>
        /// SupportByVersion Excel, 10,11,12,14,15,16
        /// This method is called by Microsoft Excel to get new data.
        /// </summary>
        /// <param name="topicCount">The RTD server must change the value of the TopicCount to the number of elements in the array returned</param>
        /// <returns>System.Array</returns>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        protected abstract object RefreshData(int topicCount);

        /// <summary>
        /// SupportByVersion Excel, 10,11,12,14,15,16
        /// Notifies a real-time data (RTD) server application that a topic is no longer in use
        /// </summary>
        /// <param name="topicID">A unique value assigned to the topic assigned by Microsoft Excel</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        protected virtual void DisconnectData(int topicID)
        {

        }

        /// <summary>
        /// SupportByVersion Excel, 10,11,12,14,15,16
        /// Determines if the real-time data server is still active. Zero or a negative number indicates failure; a positive number indicates that the server is active
        /// Default return is 1 if COMRtdServerAttribute is not set and Heartbeat is not overriden.
        /// </summary>
        /// <returns>System.Int32</returns>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        protected virtual int Heartbeat()
        {
            return null != Attribute ? Attribute.Heartbeat : 1;
        }

        /// <summary>
        /// SupportByVersion Excel, 10,11,12,14,15,16
        /// Terminates the connection to the real-time data server.
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        protected virtual void ServerTerminate()
        {

        }

        #endregion

        #region IRtdServer
       
        object Native.IRtdServer.ConnectData(int topicID, ref object strings, ref bool getNewValues)
        {
            try
            {
                lock (_thisLock)
                {
                    object result = ConnectData(topicID, Factory.WrapObject(strings, true), getNewValues);
                    return Invoker.ValidateParam(result);
                }
            }
            catch (Exception exception)
            {
                OnError(RTDMethods.ConnectData, exception);
                throw new COMRtdServerException(RTDMethods.Unknown, exception);
            }
        }

        void Native.IRtdServer.DisconnectData(int topicID)
        {
            try
            {
                lock (_thisLock)
                {
                    DisconnectData(topicID);
                }
            }
            catch (Exception exception)
            {
                OnError(RTDMethods.DisconnectData, exception);
                throw new COMRtdServerException(RTDMethods.Unknown, exception);
            }
        }

        int Native.IRtdServer.Heartbeat()
        {
            try
            {
                lock (_thisLock)
                {
                    return Heartbeat();
                }
            }
            catch (Exception exception)
            {
                OnError(RTDMethods.Heartbeat, exception);
                throw new COMRtdServerException(RTDMethods.Unknown, exception);
            }
        }

        object Native.IRtdServer.RefreshData(ref int topicCount)
        {
            try
            {
                lock (_thisLock)
                {                 
                    object result = RefreshData(topicCount);
                    return Invoker.ValidateParam(result);
                }
            }
            catch (Exception exception)
            {
                OnError(RTDMethods.RefreshData, exception);
                throw new COMRtdServerException(RTDMethods.Unknown, exception);
            }
        }

        int Native.IRtdServer.ServerStart(Native.IRTDUpdateEvent callbackObject)
        {
            try
            {
                lock (_thisLock)
                {
                    IRTDUpdateEvent arg = ReflectCallbackObject(callbackObject);
                    int result = ServerStart(arg);
                    return result;
                }
            }
            catch (Exception exception)
            {
                OnError(RTDMethods.ServerStart, exception);
                throw new COMRtdServerException(RTDMethods.Unknown, exception);
            }
        }

        void Native.IRtdServer.ServerTerminate()
        {
            try
            {
                lock (_thisLock)
                {
                    ServerTerminate();
                    Cleanup();
                }
            }
            catch (Exception exception)
            {
                OnError(RTDMethods.ServerTerminate, exception);
                throw new COMRtdServerException(RTDMethods.Unknown, exception);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Initialize the instance from ctor.
        /// The method want create the instance factory core.
        /// </summary>
        protected virtual void Initialize()
        {
            Factory = new Core();
        }

        /// <summary>
        /// Wrap native proxy into NetOffice wrapper
        /// </summary>
        /// <param name="callbackObject">given proxy from host application</param>
        /// <returns>ICOMObject proxy wrapper instance</returns>
        protected internal IRTDUpdateEvent ReflectCallbackObject(Native.IRTDUpdateEvent callbackObject)
        {
            return new IRTDUpdateEvent(Factory, null, callbackObject, IRTDUpdateEvent.LateBindingApiWrapperType);
        }

        /// <summary>
        /// Dispose the instance during IRtdServer.ServerTerminate
        /// The method want dispose all com proxies(incl. CallbackObject if its an unmodified root instance) from instance factory core.
        /// </summary>
        protected virtual void Cleanup()
        {
            try
            {
                IEnumerable<ICOMObject> roots = Factory.GetRootInstances();
                foreach (ICOMObject root in roots)
                    root.Dispose();
            }
            catch (Exception exception)
            {
                OnError(RTDMethods.Unknown, exception);
                throw new COMRtdServerException(exception);
            }
        }

        /// <summary>
        /// Called from catch handlers if an unexpected error occured
        /// </summary>
        /// <param name="methodKind">method description the error comes from</param>
        /// <param name="exception">occured exception</param>
        protected virtual void OnError(RTDMethods methodKind, Exception exception)
        {

        }

        #endregion

        #region Register

        /// <summary>
        /// Called from registration services while register
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

            MethodInfo registerMethod = null;
            RegisterFunctionAttribute registerAttribute = null;
            bool registerMethodPresent = AttributeReflector.GetRegisterAttribute(type, ref registerMethod, ref registerAttribute);
            if (null != registerAttribute && true == registerMethodPresent)
            {
                if (!TryCallDerivedRegisterMethod(registerMethod, type, InstallScope.System))
                {
                    if (!RegisterErrorHandler.RaiseStaticErrorHandlerMethod(type,
                                                                            RegisterErrorMethodKind.Register,
                                                                            new RegisterException(0)))
                        return;
                }
            }
        }

        /// <summary>
        /// Called from registration services while ungregister
        /// </summary>
        /// <param name="type">Type information for the class</param>
        [ComUnregisterFunction, Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public static void UnregisterFunction(Type type)
        {
            if (null == type)
                throw new ArgumentNullException("type");
            if (null != type.GetCustomAttribute<DontRegisterAddinAttribute>())
                return;

            UnregisterHandleProgrammable(type, InstallScope.System);
            UnregisterHandleCodebase(type, InstallScope.System);

            MethodInfo registerMethod = null;
            UnRegisterFunctionAttribute registerAttribute = null;
            bool registerMethodPresent = AttributeReflector.GetUnRegisterAttribute(type, ref registerMethod, ref registerAttribute);
            if (null != registerAttribute && true == registerMethodPresent)
            {
                if (!TryCallDerivedUnRegisterMethod(registerMethod, type, InstallScope.System))
                {
                    if (!RegisterErrorHandler.RaiseStaticErrorHandlerMethod(type,
                                                                            RegisterErrorMethodKind.UnRegister,
                                                                            new UnregisterException()))
                        return;
                }
            }
        }

        private static bool TryCallDerivedUnRegisterMethod(MethodInfo registerMethod, Type type, InstallScope scope)
        {
            try
            {
                ParameterInfo[] arguments = registerMethod.GetParameters();
                int argumentsCount = arguments.Length;
                switch (argumentsCount)
                {
                    case 0:
                        registerMethod.Invoke(null, new object[0]);
                        break;
                    case 1:
                        if (arguments[0].ParameterType.GUID == typeof(InstallScope).GUID)
                            registerMethod.Invoke(null, new object[] { scope });
                        else
                            registerMethod.Invoke(null, new object[] { type });
                        break;
                    case 2:
                        registerMethod.Invoke(null, new object[] { type, scope });
                        break;
                    default:
                        break;
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool TryCallDerivedRegisterMethod(MethodInfo registerMethod, Type type, InstallScope scope)
        {
            try
            {
                ParameterInfo[] arguments = registerMethod.GetParameters();
                int argumentsCount = arguments.Length;
                switch (argumentsCount)
                {
                    case 0:
                        registerMethod.Invoke(null, new object[0]);
                        break;
                    case 1:
                        if (arguments[0].ParameterType.GUID == typeof(InstallScope).GUID)
                            registerMethod.Invoke(null, new object[] { scope });
                        else
                            registerMethod.Invoke(null, new object[] { type });
                        break;
                    case 2:
                        registerMethod.Invoke(null, new object[] { type, scope });
                        break;
                    default:
                        break;
                }
                return true;
            }
            catch (Exception)
            {
                return false;
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
