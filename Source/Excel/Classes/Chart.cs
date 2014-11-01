using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
namespace NetOffice.ExcelApi
{

	#region Delegates

	#pragma warning disable
	public delegate void Chart_ActivateEventHandler();
	public delegate void Chart_DeactivateEventHandler();
	public delegate void Chart_ResizeEventHandler();
	public delegate void Chart_MouseDownEventHandler(Int32 Button, Int32 Shift, Int32 x, Int32 y);
	public delegate void Chart_MouseUpEventHandler(Int32 Button, Int32 Shift, Int32 x, Int32 y);
	public delegate void Chart_MouseMoveEventHandler(Int32 Button, Int32 Shift, Int32 x, Int32 y);
	public delegate void Chart_BeforeRightClickEventHandler(ref bool Cancel);
	public delegate void Chart_DragPlotEventHandler();
	public delegate void Chart_DragOverEventHandler();
	public delegate void Chart_BeforeDoubleClickEventHandler(Int32 ElementID, Int32 Arg1, Int32 Arg2, ref bool Cancel);
	public delegate void Chart_SelectEventHandler(Int32 ElementID, Int32 Arg1, Int32 Arg2);
	public delegate void Chart_SeriesChangeEventHandler(Int32 SeriesIndex, Int32 PointIndex);
	public delegate void Chart_CalculateEventHandler();
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass Chart 
	/// SupportByVersion Excel, 9,10,11,12,14,15
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194426.aspx
	///</summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class Chart : _Chart,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		ChartEvents_SinkHelper _chartEvents_SinkHelper;
	
		#endregion

		#region Type Information

        private static Type _type;
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(Chart);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Chart(Core factory, COMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Chart(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Chart(Core factory, COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Chart(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Chart(COMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		///<summary>
        ///creates a new instance of Chart 
        ///</summary>		
		public Chart():base("Excel.Chart")
		{
			
		}
		
		///<summary>
        ///creates a new instance of Chart
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public Chart(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// returns all running Excel.Chart objects from the running object table(ROT)
        /// </summary>
        /// <returns>an Excel.Chart array</returns>
		public static NetOffice.ExcelApi.Chart[] GetActiveInstances()
		{		
			NetRuntimeSystem.Collections.Generic.List<object> proxyList = NetOffice.RunningObjectTable.GetActiveProxiesFromROT("Excel","Chart");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.ExcelApi.Chart> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.ExcelApi.Chart>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.ExcelApi.Chart(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// returns a running Excel.Chart object from the running object table(ROT). the method takes the first element from the table
        /// </summary>
        /// <returns>an Excel.Chart object or null</returns>
		public static NetOffice.ExcelApi.Chart GetActiveInstance()
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Excel","Chart", false);
			if(null != proxy)
				return new NetOffice.ExcelApi.Chart(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// returns a running Excel.Chart object from the running object table(ROT).  the method takes the first element from the table
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an Excel.Chart object or null</returns>
		public static NetOffice.ExcelApi.Chart GetActiveInstance(bool throwOnError)
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Excel","Chart", throwOnError);
			if(null != proxy)
				return new NetOffice.ExcelApi.Chart(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15
		/// </summary>
		private event Chart_ActivateEventHandler _ActivateEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834456.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15)]
		public event Chart_ActivateEventHandler ActivateEvent
		{
			add
			{
				CreateEventBridge();
				_ActivateEvent += value;
			}
			remove
			{
				_ActivateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15
		/// </summary>
		private event Chart_DeactivateEventHandler _DeactivateEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838241.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15)]
		public event Chart_DeactivateEventHandler DeactivateEvent
		{
			add
			{
				CreateEventBridge();
				_DeactivateEvent += value;
			}
			remove
			{
				_DeactivateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15
		/// </summary>
		private event Chart_ResizeEventHandler _ResizeEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839406.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15)]
		public event Chart_ResizeEventHandler ResizeEvent
		{
			add
			{
				CreateEventBridge();
				_ResizeEvent += value;
			}
			remove
			{
				_ResizeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15
		/// </summary>
		private event Chart_MouseDownEventHandler _MouseDownEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822567.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15)]
		public event Chart_MouseDownEventHandler MouseDownEvent
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
		/// SupportByVersion Excel, 9,10,11,12,14,15
		/// </summary>
		private event Chart_MouseUpEventHandler _MouseUpEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197532.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15)]
		public event Chart_MouseUpEventHandler MouseUpEvent
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
		/// SupportByVersion Excel, 9,10,11,12,14,15
		/// </summary>
		private event Chart_MouseMoveEventHandler _MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837995.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15)]
		public event Chart_MouseMoveEventHandler MouseMoveEvent
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
		/// SupportByVersion Excel, 9,10,11,12,14,15
		/// </summary>
		private event Chart_BeforeRightClickEventHandler _BeforeRightClickEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839270.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15)]
		public event Chart_BeforeRightClickEventHandler BeforeRightClickEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeRightClickEvent += value;
			}
			remove
			{
				_BeforeRightClickEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15
		/// </summary>
		private event Chart_DragPlotEventHandler _DragPlotEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15)]
		public event Chart_DragPlotEventHandler DragPlotEvent
		{
			add
			{
				CreateEventBridge();
				_DragPlotEvent += value;
			}
			remove
			{
				_DragPlotEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15
		/// </summary>
		private event Chart_DragOverEventHandler _DragOverEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15)]
		public event Chart_DragOverEventHandler DragOverEvent
		{
			add
			{
				CreateEventBridge();
				_DragOverEvent += value;
			}
			remove
			{
				_DragOverEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15
		/// </summary>
		private event Chart_BeforeDoubleClickEventHandler _BeforeDoubleClickEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197223.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15)]
		public event Chart_BeforeDoubleClickEventHandler BeforeDoubleClickEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeDoubleClickEvent += value;
			}
			remove
			{
				_BeforeDoubleClickEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15
		/// </summary>
		private event Chart_SelectEventHandler _SelectEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192964.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15)]
		public event Chart_SelectEventHandler SelectEvent
		{
			add
			{
				CreateEventBridge();
				_SelectEvent += value;
			}
			remove
			{
				_SelectEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15
		/// </summary>
		private event Chart_SeriesChangeEventHandler _SeriesChangeEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834746.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15)]
		public event Chart_SeriesChangeEventHandler SeriesChangeEvent
		{
			add
			{
				CreateEventBridge();
				_SeriesChangeEvent += value;
			}
			remove
			{
				_SeriesChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15
		/// </summary>
		private event Chart_CalculateEventHandler _CalculateEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820890.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15)]
		public event Chart_CalculateEventHandler CalculateEvent
		{
			add
			{
				CreateEventBridge();
				_CalculateEvent += value;
			}
			remove
			{
				_CalculateEvent -= value;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, ChartEvents_SinkHelper.Id);


			if(ChartEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_chartEvents_SinkHelper = new ChartEvents_SinkHelper(this, _connectPoint);
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
			if( null != _chartEvents_SinkHelper)
			{
				_chartEvents_SinkHelper.Dispose();
				_chartEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}