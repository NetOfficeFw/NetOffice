using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
using NetOffice.Misc;

namespace NetOffice.ExcelApi
{

	#region Delegates

	#pragma warning disable
	public delegate void Worksheet_SelectionChangeEventHandler(NetOffice.ExcelApi.Range Target);
	public delegate void Worksheet_BeforeDoubleClickEventHandler(NetOffice.ExcelApi.Range Target, ref bool Cancel);
	public delegate void Worksheet_BeforeRightClickEventHandler(NetOffice.ExcelApi.Range Target, ref bool Cancel);
	public delegate void Worksheet_ActivateEventHandler();
	public delegate void Worksheet_DeactivateEventHandler();
	public delegate void Worksheet_CalculateEventHandler();
	public delegate void Worksheet_ChangeEventHandler(NetOffice.ExcelApi.Range Target);
	public delegate void Worksheet_FollowHyperlinkEventHandler(NetOffice.ExcelApi.Hyperlink Target);
	public delegate void Worksheet_PivotTableUpdateEventHandler(NetOffice.ExcelApi.PivotTable Target);
	public delegate void Worksheet_PivotTableAfterValueChangeEventHandler(NetOffice.ExcelApi.PivotTable TargetPivotTable, NetOffice.ExcelApi.Range TargetRange);
	public delegate void Worksheet_PivotTableBeforeAllocateChangesEventHandler(NetOffice.ExcelApi.PivotTable TargetPivotTable, Int32 ValueChangeStart, Int32 ValueChangeEnd, ref bool Cancel);
	public delegate void Worksheet_PivotTableBeforeCommitChangesEventHandler(NetOffice.ExcelApi.PivotTable TargetPivotTable, Int32 ValueChangeStart, Int32 ValueChangeEnd, ref bool Cancel);
	public delegate void Worksheet_PivotTableBeforeDiscardChangesEventHandler(NetOffice.ExcelApi.PivotTable TargetPivotTable, Int32 ValueChangeStart, Int32 ValueChangeEnd);
	public delegate void Worksheet_PivotTableChangeSyncEventHandler(NetOffice.ExcelApi.PivotTable Target);
	public delegate void Worksheet_LensGalleryRenderCompleteEventHandler();
	public delegate void Worksheet_TableUpdateEventHandler(NetOffice.ExcelApi.TableObject Target);
	public delegate void Worksheet_BeforeDeleteEventHandler();
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass Worksheet 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194464.aspx
	///</summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class Worksheet : _Worksheet,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		DocEvents_SinkHelper _docEvents_SinkHelper;
	
		#endregion

		#region Type Information

        /// <summary>
        /// Instance Type
        /// </summary>
        public override Type InstanceType
        {
            get
            {
                return LateBindingApiWrapperType;
            }
        }

        private static Type _type;
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(Worksheet);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Worksheet(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Worksheet(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Worksheet(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Worksheet(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Worksheet(ICOMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		///<summary>
        /// Creates a new instance of Worksheet 
        ///</summary>		
		public Worksheet():base("Excel.Worksheet")
		{
			
		}
		
		///<summary>
        /// Creates a new instance of Worksheet
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public Worksheet(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// Returns all running Excel.Worksheet objects from the environment/system
        /// </summary>
        /// <returns>an Excel.Worksheet array</returns>
		public static NetOffice.ExcelApi.Worksheet[] GetActiveInstances()
		{		
			IDisposableEnumeration proxyList = NetOffice.ProxyService.GetActiveInstances("Excel","Worksheet");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.ExcelApi.Worksheet> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.ExcelApi.Worksheet>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.ExcelApi.Worksheet(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// Returns a running Excel.Worksheet object from the environment/system.
        /// </summary>
        /// <returns>an Excel.Worksheet object or null</returns>
		public static NetOffice.ExcelApi.Worksheet GetActiveInstance()
		{
			object proxy  = NetOffice.ProxyService.GetActiveInstance("Excel","Worksheet", false);
			if(null != proxy)
				return new NetOffice.ExcelApi.Worksheet(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// Returns a running Excel.Worksheet object from the environment/system. 
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an Excel.Worksheet object or null</returns>
		public static NetOffice.ExcelApi.Worksheet GetActiveInstance(bool throwOnError)
		{
			object proxy  = NetOffice.ProxyService.GetActiveInstance("Excel","Worksheet", throwOnError);
			if(null != proxy)
				return new NetOffice.ExcelApi.Worksheet(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Worksheet_SelectionChangeEventHandler _SelectionChangeEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194470.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Worksheet_SelectionChangeEventHandler SelectionChangeEvent
		{
			add
			{
				CreateEventBridge();
				_SelectionChangeEvent += value;
			}
			remove
			{
				_SelectionChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Worksheet_BeforeDoubleClickEventHandler _BeforeDoubleClickEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196564.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Worksheet_BeforeDoubleClickEventHandler BeforeDoubleClickEvent
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
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Worksheet_BeforeRightClickEventHandler _BeforeRightClickEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192993.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Worksheet_BeforeRightClickEventHandler BeforeRightClickEvent
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
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Worksheet_ActivateEventHandler _ActivateEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198220.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Worksheet_ActivateEventHandler ActivateEvent
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
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Worksheet_DeactivateEventHandler _DeactivateEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197183.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Worksheet_DeactivateEventHandler DeactivateEvent
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
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Worksheet_CalculateEventHandler _CalculateEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838823.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Worksheet_CalculateEventHandler CalculateEvent
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

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Worksheet_ChangeEventHandler _ChangeEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839775.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Worksheet_ChangeEventHandler ChangeEvent
		{
			add
			{
				CreateEventBridge();
				_ChangeEvent += value;
			}
			remove
			{
				_ChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Worksheet_FollowHyperlinkEventHandler _FollowHyperlinkEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838843.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Worksheet_FollowHyperlinkEventHandler FollowHyperlinkEvent
		{
			add
			{
				CreateEventBridge();
				_FollowHyperlinkEvent += value;
			}
			remove
			{
				_FollowHyperlinkEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 10,11,12,14,15,16
		/// </summary>
		private event Worksheet_PivotTableUpdateEventHandler _PivotTableUpdateEvent;

		/// <summary>
		/// SupportByVersion Excel 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822105.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public event Worksheet_PivotTableUpdateEventHandler PivotTableUpdateEvent
		{
			add
			{
				CreateEventBridge();
				_PivotTableUpdateEvent += value;
			}
			remove
			{
				_PivotTableUpdateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 14,15,16
		/// </summary>
		private event Worksheet_PivotTableAfterValueChangeEventHandler _PivotTableAfterValueChangeEvent;

		/// <summary>
		/// SupportByVersion Excel 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193517.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public event Worksheet_PivotTableAfterValueChangeEventHandler PivotTableAfterValueChangeEvent
		{
			add
			{
				CreateEventBridge();
				_PivotTableAfterValueChangeEvent += value;
			}
			remove
			{
				_PivotTableAfterValueChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 14,15,16
		/// </summary>
		private event Worksheet_PivotTableBeforeAllocateChangesEventHandler _PivotTableBeforeAllocateChangesEvent;

		/// <summary>
		/// SupportByVersion Excel 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195070.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public event Worksheet_PivotTableBeforeAllocateChangesEventHandler PivotTableBeforeAllocateChangesEvent
		{
			add
			{
				CreateEventBridge();
				_PivotTableBeforeAllocateChangesEvent += value;
			}
			remove
			{
				_PivotTableBeforeAllocateChangesEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 14,15,16
		/// </summary>
		private event Worksheet_PivotTableBeforeCommitChangesEventHandler _PivotTableBeforeCommitChangesEvent;

		/// <summary>
		/// SupportByVersion Excel 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198138.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public event Worksheet_PivotTableBeforeCommitChangesEventHandler PivotTableBeforeCommitChangesEvent
		{
			add
			{
				CreateEventBridge();
				_PivotTableBeforeCommitChangesEvent += value;
			}
			remove
			{
				_PivotTableBeforeCommitChangesEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 14,15,16
		/// </summary>
		private event Worksheet_PivotTableBeforeDiscardChangesEventHandler _PivotTableBeforeDiscardChangesEvent;

		/// <summary>
		/// SupportByVersion Excel 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836187.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public event Worksheet_PivotTableBeforeDiscardChangesEventHandler PivotTableBeforeDiscardChangesEvent
		{
			add
			{
				CreateEventBridge();
				_PivotTableBeforeDiscardChangesEvent += value;
			}
			remove
			{
				_PivotTableBeforeDiscardChangesEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 14,15,16
		/// </summary>
		private event Worksheet_PivotTableChangeSyncEventHandler _PivotTableChangeSyncEvent;

		/// <summary>
		/// SupportByVersion Excel 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838251.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public event Worksheet_PivotTableChangeSyncEventHandler PivotTableChangeSyncEvent
		{
			add
			{
				CreateEventBridge();
				_PivotTableChangeSyncEvent += value;
			}
			remove
			{
				_PivotTableChangeSyncEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 15, 16
		/// </summary>
		private event Worksheet_LensGalleryRenderCompleteEventHandler _LensGalleryRenderCompleteEvent;

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj227603.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public event Worksheet_LensGalleryRenderCompleteEventHandler LensGalleryRenderCompleteEvent
		{
			add
			{
				CreateEventBridge();
				_LensGalleryRenderCompleteEvent += value;
			}
			remove
			{
				_LensGalleryRenderCompleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 15, 16
		/// </summary>
		private event Worksheet_TableUpdateEventHandler _TableUpdateEvent;

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229788.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public event Worksheet_TableUpdateEventHandler TableUpdateEvent
		{
			add
			{
				CreateEventBridge();
				_TableUpdateEvent += value;
			}
			remove
			{
				_TableUpdateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 15, 16
		/// </summary>
		private event Worksheet_BeforeDeleteEventHandler _BeforeDeleteEvent;

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/dn448393.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public event Worksheet_BeforeDeleteEventHandler BeforeDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeDeleteEvent += value;
			}
			remove
			{
				_BeforeDeleteEvent -= value;
			}
		}

		#endregion
       
	    #region IEventBinding Member
        
		/// <summary>
        /// Creates active sink helper
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void CreateEventBridge()
        {
			if(false == Factory.Settings.EnableEvents)
				return;
	
			if (null != _connectPoint)
				return;
	
            if (null == _activeSinkId)
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, DocEvents_SinkHelper.Id);


			if(DocEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_docEvents_SinkHelper = new DocEvents_SinkHelper(this, _connectPoint);
				return;
			} 
        }

        /// <summary>
        /// The instance use currently an event listener 
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public bool EventBridgeInitialized
        {
            get 
            {
                return (null != _connectPoint);
            }
        }

        /// <summary>
        ///  The instance has currently one or more event recipients 
        /// </summary>
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
        
        /// <summary>
        /// Target methods from its actual event recipients
        /// </summary>
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
       
        /// <summary>
        /// Returns the current count of event recipients
        /// </summary>
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
        
        /// <summary>
        /// Raise an instance event
        /// </summary>
        /// <param name="eventName">name of the event without 'Event' at the end</param>
        /// <param name="paramsArray">custom arguments for the event</param>
        /// <returns>count of called event recipients</returns>
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

        /// <summary>
        /// Stop listening events for the instance
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void DisposeEventBridge()
        {
			if( null != _docEvents_SinkHelper)
			{
				_docEvents_SinkHelper.Dispose();
				_docEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}