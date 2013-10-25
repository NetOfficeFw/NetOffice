using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
namespace NetOffice.AccessApi
{

	#region Delegates

	#pragma warning disable
	public delegate void _PageHdrFtrInReport_FormatEventHandler(ref Int16 Cancel, ref Int16 FormatCount);
	public delegate void _PageHdrFtrInReport_PrintEventHandler(ref Int16 Cancel, ref Int16 PrintCount);
	public delegate void _PageHdrFtrInReport_ClickEventHandler();
	public delegate void _PageHdrFtrInReport_DblClickEventHandler(ref Int16 Cancel);
	public delegate void _PageHdrFtrInReport_MouseDownEventHandler(ref Int16 Button, ref Int16 Shift, ref Single X, ref Single Y);
	public delegate void _PageHdrFtrInReport_MouseMoveEventHandler(ref Int16 Button, ref Int16 Shift, ref Single X, ref Single Y);
	public delegate void _PageHdrFtrInReport_MouseUpEventHandler(ref Int16 Button, ref Int16 Shift, ref Single X, ref Single Y);
	public delegate void _PageHdrFtrInReport_PaintEventHandler();
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass _PageHdrFtrInReport 
	/// SupportByVersion Access, 9,10,11,12,14,15
	///</summary>
	[SupportByVersionAttribute("Access", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class _PageHdrFtrInReport : _Section,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		_PageHdrFtrInReportEvents_SinkHelper __PageHdrFtrInReportEvents_SinkHelper;
		DispPageHdrFtrInReportEvents_SinkHelper _dispPageHdrFtrInReportEvents_SinkHelper;
	
		#endregion

		#region Type Information

        private static Type _type;
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(_PageHdrFtrInReport);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _PageHdrFtrInReport(Core factory, COMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _PageHdrFtrInReport(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _PageHdrFtrInReport(Core factory, COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _PageHdrFtrInReport(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _PageHdrFtrInReport(COMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		///<summary>
        ///creates a new instance of _PageHdrFtrInReport 
        ///</summary>		
		public _PageHdrFtrInReport():base("Access._PageHdrFtrInReport")
		{
			
		}
		
		///<summary>
        ///creates a new instance of _PageHdrFtrInReport
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public _PageHdrFtrInReport(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// returns all running Access._PageHdrFtrInReport objects from the running object table(ROT)
        /// </summary>
        /// <returns>an Access._PageHdrFtrInReport array</returns>
		public static NetOffice.AccessApi._PageHdrFtrInReport[] GetActiveInstances()
		{		
			NetRuntimeSystem.Collections.Generic.List<object> proxyList = NetOffice.RunningObjectTable.GetActiveProxiesFromROT("Access","_PageHdrFtrInReport");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.AccessApi._PageHdrFtrInReport> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.AccessApi._PageHdrFtrInReport>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.AccessApi._PageHdrFtrInReport(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// returns a running Access._PageHdrFtrInReport object from the running object table(ROT). the method takes the first element from the table
        /// </summary>
        /// <returns>an Access._PageHdrFtrInReport object or null</returns>
		public static NetOffice.AccessApi._PageHdrFtrInReport GetActiveInstance()
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Access","_PageHdrFtrInReport", false);
			if(null != proxy)
				return new NetOffice.AccessApi._PageHdrFtrInReport(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// returns a running Access._PageHdrFtrInReport object from the running object table(ROT).  the method takes the first element from the table
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an Access._PageHdrFtrInReport object or null</returns>
		public static NetOffice.AccessApi._PageHdrFtrInReport GetActiveInstance(bool throwOnError)
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Access","_PageHdrFtrInReport", throwOnError);
			if(null != proxy)
				return new NetOffice.AccessApi._PageHdrFtrInReport(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Access, 9,10,11,12,14,15
		/// </summary>
		private event _PageHdrFtrInReport_FormatEventHandler _FormatEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event _PageHdrFtrInReport_FormatEventHandler FormatEvent
		{
			add
			{
				CreateEventBridge();
				_FormatEvent += value;
			}
			remove
			{
				_FormatEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 9,10,11,12,14,15
		/// </summary>
		private event _PageHdrFtrInReport_PrintEventHandler _PrintEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event _PageHdrFtrInReport_PrintEventHandler PrintEvent
		{
			add
			{
				CreateEventBridge();
				_PrintEvent += value;
			}
			remove
			{
				_PrintEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 12,14,15
		/// </summary>
		private event _PageHdrFtrInReport_ClickEventHandler _ClickEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15
		/// </summary>
		[SupportByVersion("Access", 12,14,15)]
		public event _PageHdrFtrInReport_ClickEventHandler ClickEvent
		{
			add
			{
				CreateEventBridge();
				_ClickEvent += value;
			}
			remove
			{
				_ClickEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 12,14,15
		/// </summary>
		private event _PageHdrFtrInReport_DblClickEventHandler _DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15
		/// </summary>
		[SupportByVersion("Access", 12,14,15)]
		public event _PageHdrFtrInReport_DblClickEventHandler DblClickEvent
		{
			add
			{
				CreateEventBridge();
				_DblClickEvent += value;
			}
			remove
			{
				_DblClickEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 12,14,15
		/// </summary>
		private event _PageHdrFtrInReport_MouseDownEventHandler _MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15
		/// </summary>
		[SupportByVersion("Access", 12,14,15)]
		public event _PageHdrFtrInReport_MouseDownEventHandler MouseDownEvent
		{
			add
			{
				CreateEventBridge();
				_MouseDownEvent += value;
			}
			remove
			{
				_MouseDownEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 12,14,15
		/// </summary>
		private event _PageHdrFtrInReport_MouseMoveEventHandler _MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15
		/// </summary>
		[SupportByVersion("Access", 12,14,15)]
		public event _PageHdrFtrInReport_MouseMoveEventHandler MouseMoveEvent
		{
			add
			{
				CreateEventBridge();
				_MouseMoveEvent += value;
			}
			remove
			{
				_MouseMoveEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 12,14,15
		/// </summary>
		private event _PageHdrFtrInReport_MouseUpEventHandler _MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15
		/// </summary>
		[SupportByVersion("Access", 12,14,15)]
		public event _PageHdrFtrInReport_MouseUpEventHandler MouseUpEvent
		{
			add
			{
				CreateEventBridge();
				_MouseUpEvent += value;
			}
			remove
			{
				_MouseUpEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 12,14,15
		/// </summary>
		private event _PageHdrFtrInReport_PaintEventHandler _PaintEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15
		/// </summary>
		[SupportByVersion("Access", 12,14,15)]
		public event _PageHdrFtrInReport_PaintEventHandler PaintEvent
		{
			add
			{
				CreateEventBridge();
				_PaintEvent += value;
			}
			remove
			{
				_PaintEvent -= value;
			}
		}

		#endregion
       
	    #region IEventBinding Member
        
		/// <summary>
        /// creates active sink helper
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void CreateEventBridge()
        {
			if(false == Factory.Settings.EnableEvents)
				return;
	
			if (null != _connectPoint)
				return;
	
            if (null == _activeSinkId)
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, _PageHdrFtrInReportEvents_SinkHelper.Id,DispPageHdrFtrInReportEvents_SinkHelper.Id);


			if(_PageHdrFtrInReportEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				__PageHdrFtrInReportEvents_SinkHelper = new _PageHdrFtrInReportEvents_SinkHelper(this, _connectPoint);
				return;
			}

			if(DispPageHdrFtrInReportEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_dispPageHdrFtrInReportEvents_SinkHelper = new DispPageHdrFtrInReportEvents_SinkHelper(this, _connectPoint);
				return;
			} 
        }

        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public bool EventBridgeInitialized
        {
            get 
            {
                return (null != _connectPoint);
            }
        }
        
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public bool HasEventRecipients()       
        {
			if(null == _thisType)
				_thisType = this.GetType();
					
			foreach (NetRuntimeSystem.Reflection.EventInfo item in _thisType.GetEvents())
			{
				MulticastDelegate eventDelegate = (MulticastDelegate) _thisType.GetType().GetField(item.Name, 
																			NetRuntimeSystem.Reflection.BindingFlags.NonPublic |
																			NetRuntimeSystem.Reflection.BindingFlags.Instance).GetValue(this);
					
				if( (null != eventDelegate) && (eventDelegate.GetInvocationList().Length > 0) )
					return false;
			}
				
			return false;
        }
        
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Delegate[] GetEventRecipients(string eventName)
        {
			if(null == _thisType)
				_thisType = this.GetType();
             
            MulticastDelegate eventDelegate = (MulticastDelegate)_thisType.GetField(
                                                "_" + eventName + "Event",
                                                NetRuntimeSystem.Reflection.BindingFlags.Instance |
                                                NetRuntimeSystem.Reflection.BindingFlags.NonPublic).GetValue(this);

            if (null != eventDelegate)
            {
                Delegate[] delegates = eventDelegate.GetInvocationList();
                return delegates;
            }
            else
                return new Delegate[0];
        }

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public int GetCountOfEventRecipients(string eventName)
        {
			if(null == _thisType)
				_thisType = this.GetType();
             
            MulticastDelegate eventDelegate = (MulticastDelegate)_thisType.GetField(
                                                "_" + eventName + "Event",
                                                NetRuntimeSystem.Reflection.BindingFlags.Instance |
                                                NetRuntimeSystem.Reflection.BindingFlags.NonPublic).GetValue(this);

            if (null != eventDelegate)
            {
                Delegate[] delegates = eventDelegate.GetInvocationList();
                return delegates.Length;
            }
            else
                return 0;
        }

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public int RaiseCustomEvent(string eventName, ref object[] paramsArray)
		{
			if(null == _thisType)
				_thisType = this.GetType();
             
            MulticastDelegate eventDelegate = (MulticastDelegate)_thisType.GetField(
                                                "_" + eventName + "Event",
                                                NetRuntimeSystem.Reflection.BindingFlags.Instance |
                                                NetRuntimeSystem.Reflection.BindingFlags.NonPublic).GetValue(this);

            if (null != eventDelegate)
            {
                Delegate[] delegates = eventDelegate.GetInvocationList();
                foreach (var item in delegates)
                {
                    try
                    {
                        item.Method.Invoke(item.Target, paramsArray);
                    }
                    catch (NetRuntimeSystem.Exception exception)
                    {
                        Factory.Console.WriteException(exception);
                    }
                }
                return delegates.Length;
            }
            else
                return 0;
		}

        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void DisposeEventBridge()
        {
			if( null != __PageHdrFtrInReportEvents_SinkHelper)
			{
				__PageHdrFtrInReportEvents_SinkHelper.Dispose();
				__PageHdrFtrInReportEvents_SinkHelper = null;
			}
			if( null != _dispPageHdrFtrInReportEvents_SinkHelper)
			{
				_dispPageHdrFtrInReportEvents_SinkHelper.Dispose();
				_dispPageHdrFtrInReportEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}