using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Application 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821758.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass), ComProgId("Access.Application"), ModuleProvider(typeof(GlobalHelperModules.GlobalModule))]
    public class Application : _Application, ICloneable<Application>
    {
		#pragma warning disable

		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;

        #endregion

        #region Type Information

        /// <summary>
        /// Instance Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type InstanceType
        {
            get
            {
                return LateBindingApiWrapperType;
            }
        }
        
        /// <summary>
        /// Type Cache
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(Application);
                return _type;
            }
        }
        
        #endregion
        		
		#region Ctor

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Application(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			_callQuitInDispose = true;
			GlobalHelperModules.GlobalModule.Instance = this;
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Application(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			_callQuitInDispose = true;
			GlobalHelperModules.GlobalModule.Instance = this;
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Application(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			_callQuitInDispose = true;
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Application(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			_callQuitInDispose = true;
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Application(ICOMObject replacedObject) : base(replacedObject)
		{
			_callQuitInDispose = true;
		}
		
		/// <summary>
        /// Creates a new instance of Application
        /// </summary>
        ///<param name="progId">registered ProgID</param>
		public Application(string progId):base(progId)
		{
			_callQuitInDispose = true;
			GlobalHelperModules.GlobalModule.Instance = this;
		}

        /// <summary>
        /// Creates a new instance of Application 
        /// </summary>		
        public Application() : this(false)
        {

        }

        /// <summary>
        /// Creates a new instance of Application 
        /// <param name="enableProxyService">try to get a running application first before create a new application</param>
        /// </summary>		
        public Application(bool enableProxyService = false) : base()
        {
            if (enableProxyService)
            {
                Factory = Core.Default;
                object proxy = Running.ProxyService.GetActiveInstance("Access", "Application", false);
                if (null != proxy)
                {
                    CreateFromProxy(proxy, true);
                    FromProxyService = true;
                }
                else
                {
                    CreateFromProgId("Access.Application", true);
                }
            }
            else
            {
                CreateFromProgId("Access.Application", true);
            }

            OnCreate();
            _callQuitInDispose = true;
            GlobalHelperModules.GlobalModule.Instance = this;
        }

        /// <summary>
		/// NetOffice method: dispose instance and all child instances
		/// </summary>
		/// <param name="disposeEventBinding">dispose event exported proxies with one or more event recipients</param>
		[Category("NetOffice"), CoreOverridden]
		public override void Dispose(bool disposeEventBinding)
		{
			if(this.Equals(GlobalHelperModules.GlobalModule.Instance))
				 GlobalHelperModules.GlobalModule.Instance = null;	
			base.Dispose(disposeEventBinding);
		}

		/// <summary>
		/// NetOffice method: dispose instance and all child instances
		/// </summary>
		[Category("NetOffice"), CoreOverridden]
		public override void Dispose()
		{
			if(this.Equals(GlobalHelperModules.GlobalModule.Instance))
				 GlobalHelperModules.GlobalModule.Instance = null;
			base.Dispose();
		}

        #endregion

        #region Properties

        /// <summary>
        /// Instance is created from an already running application
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public bool FromProxyService { get; private set; }

        #endregion

        #region Static CoClass Methods

        /// <summary>
        /// Returns all running Access.Application instances from the environment/system
        /// </summary>
        /// <returns>Access.Application sequence</returns>
        public static IDisposableSequence<Application> GetActiveInstances()
        {
            return Running.ProxyService.GetActiveInstances<Application>("Access", "Application");
        }

        /// <summary>
        /// Returns a running Access.Application instance from the environment/system
        /// </summary>
        /// <param name="throwExceptionIfNotFound">throw exception if unable to find an instance</param>
        /// <returns>Access.Application instance or null</returns>
        public static Application GetActiveInstance(bool throwExceptionIfNotFound = false)
        {
            return Running.ProxyService.GetActiveInstance<Application>("Access", "Application", throwExceptionIfNotFound);
        }

        #endregion

        #region Events

        #endregion

        #region ICloneable<Application>

        /// <summary>
        /// Creates a new Application that is a copy of the current instance
        /// </summary>
        /// <returns>A new Application that is a copy of this instance</returns>
        /// <exception cref="CloneException">An unexpected error occured. See inner exception(s) for details.</exception>
        public new virtual Application Clone()
        {
            return base.Clone() as Application;
        }

        #endregion

        #pragma warning restore
    }
}