﻿using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Chart_ActivateEventHandler();
	public delegate void Chart_DeactivateEventHandler();
	public delegate void Chart_ResizeEventHandler();
	public delegate void Chart_MouseDownEventHandler(Int32 button, Int32 shift, Int32 x, Int32 y);
	public delegate void Chart_MouseUpEventHandler(Int32 button, Int32 shift, Int32 x, Int32 y);
	public delegate void Chart_MouseMoveEventHandler(Int32 button, Int32 shift, Int32 x, Int32 y);
	public delegate void Chart_BeforeRightClickEventHandler(ref bool cancel);
	public delegate void Chart_DragPlotEventHandler();
	public delegate void Chart_DragOverEventHandler();
	public delegate void Chart_BeforeDoubleClickEventHandler(Int32 elementID, Int32 arg1, Int32 arg2, ref bool cancel);
	public delegate void Chart_SelectEventHandler(Int32 elementID, Int32 arg1, Int32 arg2);
	public delegate void Chart_SeriesChangeEventHandler(Int32 seriesIndex, Int32 pointIndex);
	public delegate void Chart_CalculateEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Chart 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Chart(object)"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[EventSink(typeof(Events.ChartEvents_SinkHelper))]
    [ComEventInterface(typeof(Events.ChartEvents))]
    public class Chart : _Chart, IEventBinding
	{
		#pragma warning disable

		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;
        private Events.ChartEvents_SinkHelper _chartEvents_SinkHelper;
	
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
                    _type = typeof(Chart);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Chart(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Chart(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Chart(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Chart(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Chart(ICOMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		/// <summary>
        /// Creates a new instance of Chart 
        /// </summary>		
		public Chart():base("Excel.Chart")
		{
			
		}
		
		/// <summary>
        /// Creates a new instance of Chart
        /// </summary>
        ///<param name="progId">registered ProgID</param>
		public Chart(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Chart_ActivateEventHandler _ActivateEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Chart.Activate(even)"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Chart_DeactivateEventHandler _DeactivateEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Chart.Deactivate"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Chart_ResizeEventHandler _ResizeEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Chart.Resize"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Chart_MouseDownEventHandler _MouseDownEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Chart.MouseDown"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Chart_MouseUpEventHandler _MouseUpEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Chart.MouseUp"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Chart_MouseMoveEventHandler _MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Chart.MouseMove"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Chart_BeforeRightClickEventHandler _BeforeRightClickEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Chart.BeforeRightClick"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Chart_DragPlotEventHandler _DragPlotEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Chart_DragOverEventHandler _DragOverEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Chart_BeforeDoubleClickEventHandler _BeforeDoubleClickEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Chart.BeforeDoubleClick"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Chart_SelectEventHandler _SelectEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Chart.Select(even)"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Chart_SeriesChangeEventHandler _SeriesChangeEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Chart.SeriesChange"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Chart_CalculateEventHandler _CalculateEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Chart.Calculate"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
       
	    #region IEventBinding
        
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, Events.ChartEvents_SinkHelper.Id);


			if(Events.ChartEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_chartEvents_SinkHelper = new Events.ChartEvents_SinkHelper(this, _connectPoint);
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
        /// Instance has one or more event recipients
        /// </summary>
        /// <returns>true if one or more event is active, otherwise false</returns>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public bool HasEventRecipients()       
        {
            return NetOffice.Events.CoClassEventReflector.HasEventRecipients(this, LateBindingApiWrapperType);            
        }

        /// <summary>
        /// Instance has one or more event recipients
        /// </summary>
        /// <param name="eventName">name of the event</param>
        /// <returns></returns>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public bool HasEventRecipients(string eventName)
        {
            return NetOffice.Events.CoClassEventReflector.HasEventRecipients(this, LateBindingApiWrapperType, eventName);
        }

        /// <summary>
        /// Target methods from its actual event recipients
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Delegate[] GetEventRecipients(string eventName)
        {
            return NetOffice.Events.CoClassEventReflector.GetEventRecipients(this, LateBindingApiWrapperType, eventName);
        }
       
        /// <summary>
        /// Returns the current count of event recipients
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public int GetCountOfEventRecipients(string eventName)
        {
            return NetOffice.Events.CoClassEventReflector.GetCountOfEventRecipients(this, LateBindingApiWrapperType, eventName);       
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
            return NetOffice.Events.CoClassEventReflector.RaiseCustomEvent(this, LateBindingApiWrapperType, eventName, ref paramsArray);
		}
        /// <summary>
        /// Stop listening events for the instance
        /// </summary>
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

