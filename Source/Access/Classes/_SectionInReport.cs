using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
namespace NetOffice.AccessApi
{

	#region Delegates

	#pragma warning disable
	public delegate void _SectionInReport_FormatEventHandler(ref Int16 Cancel, ref Int16 FormatCount);
	public delegate void _SectionInReport_PrintEventHandler(ref Int16 Cancel, ref Int16 PrintCount);
	public delegate void _SectionInReport_RetreatEventHandler();
	public delegate void _SectionInReport_ClickEventHandler();
	public delegate void _SectionInReport_DblClickEventHandler(ref Int16 Cancel);
	public delegate void _SectionInReport_MouseDownEventHandler(ref Int16 Button, ref Int16 Shift, ref Single X, ref Single Y);
	public delegate void _SectionInReport_MouseMoveEventHandler(ref Int16 Button, ref Int16 Shift, ref Single X, ref Single Y);
	public delegate void _SectionInReport_MouseUpEventHandler(ref Int16 Button, ref Int16 Shift, ref Single X, ref Single Y);
	public delegate void _SectionInReport_PaintEventHandler();
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass _SectionInReport 
	/// SupportByVersion Access, 9,10,11,12,14,15
	///</summary>
	[SupportByVersionAttribute("Access", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class _SectionInReport : _Section,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		_SectionInReportEvents_SinkHelper __SectionInReportEvents_SinkHelper;
		DispSectionInReportEvents_SinkHelper _dispSectionInReportEvents_SinkHelper;
	
		#endregion

		#region Type Information

        private static Type _type;
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(_SectionInReport);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _SectionInReport(Core factory, COMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _SectionInReport(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _SectionInReport(Core factory, COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _SectionInReport(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _SectionInReport(COMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		///<summary>
        ///creates a new instance of _SectionInReport 
        ///</summary>		
		public _SectionInReport():base("Access._SectionInReport")
		{
			
		}
		
		///<summary>
        ///creates a new instance of _SectionInReport
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public _SectionInReport(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// returns all running Access._SectionInReport objects from the running object table(ROT)
        /// </summary>
        /// <returns>an Access._SectionInReport array</returns>
		public static NetOffice.AccessApi._SectionInReport[] GetActiveInstances()
		{		
			NetRuntimeSystem.Collections.Generic.List<object> proxyList = NetOffice.RunningObjectTable.GetActiveProxiesFromROT("Access","_SectionInReport");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.AccessApi._SectionInReport> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.AccessApi._SectionInReport>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.AccessApi._SectionInReport(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// returns a running Access._SectionInReport object from the running object table(ROT). the method takes the first element from the table
        /// </summary>
        /// <returns>an Access._SectionInReport object or null</returns>
		public static NetOffice.AccessApi._SectionInReport GetActiveInstance()
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Access","_SectionInReport", false);
			if(null != proxy)
				return new NetOffice.AccessApi._SectionInReport(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// returns a running Access._SectionInReport object from the running object table(ROT).  the method takes the first element from the table
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an Access._SectionInReport object or null</returns>
		public static NetOffice.AccessApi._SectionInReport GetActiveInstance(bool throwOnError)
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Access","_SectionInReport", throwOnError);
			if(null != proxy)
				return new NetOffice.AccessApi._SectionInReport(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Access, 9,10,11,12,14,15
		/// </summary>
		private event _SectionInReport_FormatEventHandler _FormatEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event _SectionInReport_FormatEventHandler FormatEvent
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
		private event _SectionInReport_PrintEventHandler _PrintEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event _SectionInReport_PrintEventHandler PrintEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15
		/// </summary>
		private event _SectionInReport_RetreatEventHandler _RetreatEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event _SectionInReport_RetreatEventHandler RetreatEvent
		{
			add
			{
				CreateEventBridge();
				_RetreatEvent += value;
			}
			remove
			{
				_RetreatEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Access, 12,14,15
		/// </summary>
		private event _SectionInReport_ClickEventHandler _ClickEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15
		/// </summary>
		[SupportByVersion("Access", 12,14,15)]
		public event _SectionInReport_ClickEventHandler ClickEvent
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
		private event _SectionInReport_DblClickEventHandler _DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15
		/// </summary>
		[SupportByVersion("Access", 12,14,15)]
		public event _SectionInReport_DblClickEventHandler DblClickEvent
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
		private event _SectionInReport_MouseDownEventHandler _MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15
		/// </summary>
		[SupportByVersion("Access", 12,14,15)]
		public event _SectionInReport_MouseDownEventHandler MouseDownEvent
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
		private event _SectionInReport_MouseMoveEventHandler _MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15
		/// </summary>
		[SupportByVersion("Access", 12,14,15)]
		public event _SectionInReport_MouseMoveEventHandler MouseMoveEvent
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
		private event _SectionInReport_MouseUpEventHandler _MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15
		/// </summary>
		[SupportByVersion("Access", 12,14,15)]
		public event _SectionInReport_MouseUpEventHandler MouseUpEvent
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
		private event _SectionInReport_PaintEventHandler _PaintEvent;

		/// <summary>
		/// SupportByVersion Access 12 14 15
		/// </summary>
		[SupportByVersion("Access", 12,14,15)]
		public event _SectionInReport_PaintEventHandler PaintEvent
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, _SectionInReportEvents_SinkHelper.Id,DispSectionInReportEvents_SinkHelper.Id);


			if(_SectionInReportEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				__SectionInReportEvents_SinkHelper = new _SectionInReportEvents_SinkHelper(this, _connectPoint);
				return;
			}

			if(DispSectionInReportEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_dispSectionInReportEvents_SinkHelper = new DispSectionInReportEvents_SinkHelper(this, _connectPoint);
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
			if( null != __SectionInReportEvents_SinkHelper)
			{
				__SectionInReportEvents_SinkHelper.Dispose();
				__SectionInReportEvents_SinkHelper = null;
			}
			if( null != _dispSectionInReportEvents_SinkHelper)
			{
				_dispSectionInReportEvents_SinkHelper.Dispose();
				_dispSectionInReportEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}