using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
namespace NetOffice.AccessApi
{

	#region Delegates

	#pragma warning disable
	public delegate void Rectangle_ClickEventHandler();
	public delegate void Rectangle_DblClickEventHandler(ref Int16 Cancel);
	public delegate void Rectangle_MouseDownEventHandler(ref Int16 Button, ref Int16 Shift, ref Single X, ref Single Y);
	public delegate void Rectangle_MouseMoveEventHandler(ref Int16 Button, ref Int16 Shift, ref Single X, ref Single Y);
	public delegate void Rectangle_MouseUpEventHandler(ref Int16 Button, ref Int16 Shift, ref Single X, ref Single Y);
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass Rectangle 
	/// SupportByVersion Access, 9,10,11,12,14,15
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836237.aspx
	///</summary>
	[SupportByVersionAttribute("Access", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class Rectangle : _Rectangle,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		_RectangleEvents_SinkHelper __RectangleEvents_SinkHelper;
		DispRectangleEvents_SinkHelper _dispRectangleEvents_SinkHelper;
	
		#endregion

		#region Type Information

        private static Type _type;
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(Rectangle);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Rectangle(Core factory, COMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Rectangle(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Rectangle(Core factory, COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Rectangle(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Rectangle(COMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		///<summary>
        ///creates a new instance of Rectangle 
        ///</summary>		
		public Rectangle():base("Access.Rectangle")
		{
			
		}
		
		///<summary>
        ///creates a new instance of Rectangle
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public Rectangle(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// returns all running Access.Rectangle objects from the running object table(ROT)
        /// </summary>
        /// <returns>an Access.Rectangle array</returns>
		public static NetOffice.AccessApi.Rectangle[] GetActiveInstances()
		{		
			NetRuntimeSystem.Collections.Generic.List<object> proxyList = NetOffice.RunningObjectTable.GetActiveProxiesFromROT("Access","Rectangle");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.AccessApi.Rectangle> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.AccessApi.Rectangle>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.AccessApi.Rectangle(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// returns a running Access.Rectangle object from the running object table(ROT). the method takes the first element from the table
        /// </summary>
        /// <returns>an Access.Rectangle object or null</returns>
		public static NetOffice.AccessApi.Rectangle GetActiveInstance()
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Access","Rectangle", false);
			if(null != proxy)
				return new NetOffice.AccessApi.Rectangle(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// returns a running Access.Rectangle object from the running object table(ROT).  the method takes the first element from the table
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an Access.Rectangle object or null</returns>
		public static NetOffice.AccessApi.Rectangle GetActiveInstance(bool throwOnError)
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Access","Rectangle", throwOnError);
			if(null != proxy)
				return new NetOffice.AccessApi.Rectangle(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Access, 9,10,11,12,14,15
		/// </summary>
		private event Rectangle_ClickEventHandler _ClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197703.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event Rectangle_ClickEventHandler ClickEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15
		/// </summary>
		private event Rectangle_DblClickEventHandler _DblClickEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834418.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event Rectangle_DblClickEventHandler DblClickEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15
		/// </summary>
		private event Rectangle_MouseDownEventHandler _MouseDownEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845342.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event Rectangle_MouseDownEventHandler MouseDownEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15
		/// </summary>
		private event Rectangle_MouseMoveEventHandler _MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192713.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event Rectangle_MouseMoveEventHandler MouseMoveEvent
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
		/// SupportByVersion Access, 9,10,11,12,14,15
		/// </summary>
		private event Rectangle_MouseUpEventHandler _MouseUpEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834363.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15)]
		public event Rectangle_MouseUpEventHandler MouseUpEvent
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, _RectangleEvents_SinkHelper.Id,DispRectangleEvents_SinkHelper.Id);


			if(_RectangleEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				__RectangleEvents_SinkHelper = new _RectangleEvents_SinkHelper(this, _connectPoint);
				return;
			}

			if(DispRectangleEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_dispRectangleEvents_SinkHelper = new DispRectangleEvents_SinkHelper(this, _connectPoint);
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
			if( null != __RectangleEvents_SinkHelper)
			{
				__RectangleEvents_SinkHelper.Dispose();
				__RectangleEvents_SinkHelper = null;
			}
			if( null != _dispRectangleEvents_SinkHelper)
			{
				_dispRectangleEvents_SinkHelper.Dispose();
				_dispRectangleEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}