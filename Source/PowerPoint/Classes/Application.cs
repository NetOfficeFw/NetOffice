using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
using NetOffice.Misc;

namespace NetOffice.PowerPointApi
{

	#region Delegates

	#pragma warning disable
	public delegate void Application_WindowSelectionChangeEventHandler(NetOffice.PowerPointApi.Selection Sel);
	public delegate void Application_WindowBeforeRightClickEventHandler(NetOffice.PowerPointApi.Selection Sel, ref bool Cancel);
	public delegate void Application_WindowBeforeDoubleClickEventHandler(NetOffice.PowerPointApi.Selection Sel, ref bool Cancel);
	public delegate void Application_PresentationCloseEventHandler(NetOffice.PowerPointApi.Presentation Pres);
	public delegate void Application_PresentationSaveEventHandler(NetOffice.PowerPointApi.Presentation Pres);
	public delegate void Application_PresentationOpenEventHandler(NetOffice.PowerPointApi.Presentation Pres);
	public delegate void Application_NewPresentationEventHandler(NetOffice.PowerPointApi.Presentation Pres);
	public delegate void Application_PresentationNewSlideEventHandler(NetOffice.PowerPointApi.Slide Sld);
	public delegate void Application_WindowActivateEventHandler(NetOffice.PowerPointApi.Presentation Pres, NetOffice.PowerPointApi.DocumentWindow Wn);
	public delegate void Application_WindowDeactivateEventHandler(NetOffice.PowerPointApi.Presentation Pres, NetOffice.PowerPointApi.DocumentWindow Wn);
	public delegate void Application_SlideShowBeginEventHandler(NetOffice.PowerPointApi.SlideShowWindow Wn);
	public delegate void Application_SlideShowNextBuildEventHandler(NetOffice.PowerPointApi.SlideShowWindow Wn);
	public delegate void Application_SlideShowNextSlideEventHandler(NetOffice.PowerPointApi.SlideShowWindow Wn);
	public delegate void Application_SlideShowEndEventHandler(NetOffice.PowerPointApi.Presentation Pres);
	public delegate void Application_PresentationPrintEventHandler(NetOffice.PowerPointApi.Presentation Pres);
	public delegate void Application_SlideSelectionChangedEventHandler(NetOffice.PowerPointApi.SlideRange SldRange);
	public delegate void Application_ColorSchemeChangedEventHandler(NetOffice.PowerPointApi.SlideRange SldRange);
	public delegate void Application_PresentationBeforeSaveEventHandler(NetOffice.PowerPointApi.Presentation Pres, ref bool Cancel);
	public delegate void Application_SlideShowNextClickEventHandler(NetOffice.PowerPointApi.SlideShowWindow Wn, NetOffice.PowerPointApi.Effect nEffect);
	public delegate void Application_AfterNewPresentationEventHandler(NetOffice.PowerPointApi.Presentation Pres);
	public delegate void Application_AfterPresentationOpenEventHandler(NetOffice.PowerPointApi.Presentation Pres);
	public delegate void Application_PresentationSyncEventHandler(NetOffice.PowerPointApi.Presentation Pres, NetOffice.OfficeApi.Enums.MsoSyncEventType SyncEventType);
	public delegate void Application_SlideShowOnNextEventHandler(NetOffice.PowerPointApi.SlideShowWindow Wn);
	public delegate void Application_SlideShowOnPreviousEventHandler(NetOffice.PowerPointApi.SlideShowWindow Wn);
	public delegate void Application_PresentationBeforeCloseEventHandler(NetOffice.PowerPointApi.Presentation Pres, ref bool Cancel);
	public delegate void Application_ProtectedViewWindowOpenEventHandler(NetOffice.PowerPointApi.ProtectedViewWindow ProtViewWindow);
	public delegate void Application_ProtectedViewWindowBeforeEditEventHandler(NetOffice.PowerPointApi.ProtectedViewWindow ProtViewWindow, ref bool Cancel);
	public delegate void Application_ProtectedViewWindowBeforeCloseEventHandler(NetOffice.PowerPointApi.ProtectedViewWindow ProtViewWindow, NetOffice.PowerPointApi.Enums.PpProtectedViewCloseReason ProtectedViewCloseReason, ref bool Cancel);
	public delegate void Application_ProtectedViewWindowActivateEventHandler(NetOffice.PowerPointApi.ProtectedViewWindow ProtViewWindow);
	public delegate void Application_ProtectedViewWindowDeactivateEventHandler(NetOffice.PowerPointApi.ProtectedViewWindow ProtViewWindow);
	public delegate void Application_PresentationCloseFinalEventHandler(NetOffice.PowerPointApi.Presentation Pres);
	public delegate void Application_AfterDragDropOnSlideEventHandler(NetOffice.PowerPointApi.Slide Sld, Single X, Single Y);
	public delegate void Application_AfterShapeSizeChangeEventHandler(NetOffice.PowerPointApi.Shape shp);
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass Application 
	/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745704.aspx
	///</summary>
	[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class Application : _Application,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		EApplication_SinkHelper _eApplication_SinkHelper;
	
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
                    _type = typeof(Application);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

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
		
		///<summary>
        /// Creates a new instance of Application 
        ///</summary>		
		public Application():base("PowerPoint.Application")
		{
			_callQuitInDispose = true;
			GlobalHelperModules.GlobalModule.Instance = this;
		}
		
		///<summary>
        /// Creates a new instance of Application
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public Application(string progId):base(progId)
		{
			_callQuitInDispose = true;
			GlobalHelperModules.GlobalModule.Instance = this;
		}
		
/// <summary>
		/// NetOffice method: dispose instance and all child instances
		/// </summary>
		/// <param name="disposeEventBinding">dispose event exported proxies with one or more event recipients</param>
		public override void Dispose(bool disposeEventBinding)
		{
			if(this.Equals(GlobalHelperModules.GlobalModule.Instance))
				 GlobalHelperModules.GlobalModule.Instance = null;	
			base.Dispose(disposeEventBinding);
		}

		/// <summary>
		/// NetOffice method: dispose instance and all child instances
		/// </summary>
		public override void Dispose()
		{
			if(this.Equals(GlobalHelperModules.GlobalModule.Instance))
				 GlobalHelperModules.GlobalModule.Instance = null;
			base.Dispose();
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// Returns all running PowerPoint.Application objects from the environment/system
        /// </summary>
        /// <returns>an PowerPoint.Application array</returns>
		public static NetOffice.PowerPointApi.Application[] GetActiveInstances()
		{		
			IDisposableEnumeration proxyList = NetOffice.ProxyService.GetActiveInstances("PowerPoint","Application");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.PowerPointApi.Application> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.PowerPointApi.Application>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.PowerPointApi.Application(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// Returns a running PowerPoint.Application object from the environment/system.
        /// </summary>
        /// <returns>an PowerPoint.Application object or null</returns>
		public static NetOffice.PowerPointApi.Application GetActiveInstance()
		{
			object proxy  = NetOffice.ProxyService.GetActiveInstance("PowerPoint","Application", false);
			if(null != proxy)
				return new NetOffice.PowerPointApi.Application(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// Returns a running PowerPoint.Application object from the environment/system. 
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an PowerPoint.Application object or null</returns>
		public static NetOffice.PowerPointApi.Application GetActiveInstance(bool throwOnError)
		{
			object proxy  = NetOffice.ProxyService.GetActiveInstance("PowerPoint","Application", throwOnError);
			if(null != proxy)
				return new NetOffice.PowerPointApi.Application(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_WindowSelectionChangeEventHandler _WindowSelectionChangeEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff743918.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public event Application_WindowSelectionChangeEventHandler WindowSelectionChangeEvent
		{
			add
			{
				CreateEventBridge();
				_WindowSelectionChangeEvent += value;
			}
			remove
			{
				_WindowSelectionChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_WindowBeforeRightClickEventHandler _WindowBeforeRightClickEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746559.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public event Application_WindowBeforeRightClickEventHandler WindowBeforeRightClickEvent
		{
			add
			{
				CreateEventBridge();
				_WindowBeforeRightClickEvent += value;
			}
			remove
			{
				_WindowBeforeRightClickEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_WindowBeforeDoubleClickEventHandler _WindowBeforeDoubleClickEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745746.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public event Application_WindowBeforeDoubleClickEventHandler WindowBeforeDoubleClickEvent
		{
			add
			{
				CreateEventBridge();
				_WindowBeforeDoubleClickEvent += value;
			}
			remove
			{
				_WindowBeforeDoubleClickEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_PresentationCloseEventHandler _PresentationCloseEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744678.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public event Application_PresentationCloseEventHandler PresentationCloseEvent
		{
			add
			{
				CreateEventBridge();
				_PresentationCloseEvent += value;
			}
			remove
			{
				_PresentationCloseEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_PresentationSaveEventHandler _PresentationSaveEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744230.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public event Application_PresentationSaveEventHandler PresentationSaveEvent
		{
			add
			{
				CreateEventBridge();
				_PresentationSaveEvent += value;
			}
			remove
			{
				_PresentationSaveEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_PresentationOpenEventHandler _PresentationOpenEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744100.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public event Application_PresentationOpenEventHandler PresentationOpenEvent
		{
			add
			{
				CreateEventBridge();
				_PresentationOpenEvent += value;
			}
			remove
			{
				_PresentationOpenEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_NewPresentationEventHandler _NewPresentationEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745073.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public event Application_NewPresentationEventHandler NewPresentationEvent
		{
			add
			{
				CreateEventBridge();
				_NewPresentationEvent += value;
			}
			remove
			{
				_NewPresentationEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_PresentationNewSlideEventHandler _PresentationNewSlideEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746597.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public event Application_PresentationNewSlideEventHandler PresentationNewSlideEvent
		{
			add
			{
				CreateEventBridge();
				_PresentationNewSlideEvent += value;
			}
			remove
			{
				_PresentationNewSlideEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_WindowActivateEventHandler _WindowActivateEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff743995.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public event Application_WindowActivateEventHandler WindowActivateEvent
		{
			add
			{
				CreateEventBridge();
				_WindowActivateEvent += value;
			}
			remove
			{
				_WindowActivateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_WindowDeactivateEventHandler _WindowDeactivateEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745519.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public event Application_WindowDeactivateEventHandler WindowDeactivateEvent
		{
			add
			{
				CreateEventBridge();
				_WindowDeactivateEvent += value;
			}
			remove
			{
				_WindowDeactivateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_SlideShowBeginEventHandler _SlideShowBeginEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746741.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public event Application_SlideShowBeginEventHandler SlideShowBeginEvent
		{
			add
			{
				CreateEventBridge();
				_SlideShowBeginEvent += value;
			}
			remove
			{
				_SlideShowBeginEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_SlideShowNextBuildEventHandler _SlideShowNextBuildEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745070.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public event Application_SlideShowNextBuildEventHandler SlideShowNextBuildEvent
		{
			add
			{
				CreateEventBridge();
				_SlideShowNextBuildEvent += value;
			}
			remove
			{
				_SlideShowNextBuildEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_SlideShowNextSlideEventHandler _SlideShowNextSlideEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745863.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public event Application_SlideShowNextSlideEventHandler SlideShowNextSlideEvent
		{
			add
			{
				CreateEventBridge();
				_SlideShowNextSlideEvent += value;
			}
			remove
			{
				_SlideShowNextSlideEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_SlideShowEndEventHandler _SlideShowEndEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746536.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public event Application_SlideShowEndEventHandler SlideShowEndEvent
		{
			add
			{
				CreateEventBridge();
				_SlideShowEndEvent += value;
			}
			remove
			{
				_SlideShowEndEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_PresentationPrintEventHandler _PresentationPrintEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744696.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public event Application_PresentationPrintEventHandler PresentationPrintEvent
		{
			add
			{
				CreateEventBridge();
				_PresentationPrintEvent += value;
			}
			remove
			{
				_PresentationPrintEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 10,11,12,14,15,16
		/// </summary>
		private event Application_SlideSelectionChangedEventHandler _SlideSelectionChangedEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745869.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public event Application_SlideSelectionChangedEventHandler SlideSelectionChangedEvent
		{
			add
			{
				CreateEventBridge();
				_SlideSelectionChangedEvent += value;
			}
			remove
			{
				_SlideSelectionChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 10,11,12,14,15,16
		/// </summary>
		private event Application_ColorSchemeChangedEventHandler _ColorSchemeChangedEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745549.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public event Application_ColorSchemeChangedEventHandler ColorSchemeChangedEvent
		{
			add
			{
				CreateEventBridge();
				_ColorSchemeChangedEvent += value;
			}
			remove
			{
				_ColorSchemeChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 10,11,12,14,15,16
		/// </summary>
		private event Application_PresentationBeforeSaveEventHandler _PresentationBeforeSaveEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744682.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public event Application_PresentationBeforeSaveEventHandler PresentationBeforeSaveEvent
		{
			add
			{
				CreateEventBridge();
				_PresentationBeforeSaveEvent += value;
			}
			remove
			{
				_PresentationBeforeSaveEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 10,11,12,14,15,16
		/// </summary>
		private event Application_SlideShowNextClickEventHandler _SlideShowNextClickEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745682.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public event Application_SlideShowNextClickEventHandler SlideShowNextClickEvent
		{
			add
			{
				CreateEventBridge();
				_SlideShowNextClickEvent += value;
			}
			remove
			{
				_SlideShowNextClickEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 11,12,14,15,16
		/// </summary>
		private event Application_AfterNewPresentationEventHandler _AfterNewPresentationEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746421.aspx </remarks>
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
		public event Application_AfterNewPresentationEventHandler AfterNewPresentationEvent
		{
			add
			{
				CreateEventBridge();
				_AfterNewPresentationEvent += value;
			}
			remove
			{
				_AfterNewPresentationEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 11,12,14,15,16
		/// </summary>
		private event Application_AfterPresentationOpenEventHandler _AfterPresentationOpenEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744659.aspx </remarks>
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
		public event Application_AfterPresentationOpenEventHandler AfterPresentationOpenEvent
		{
			add
			{
				CreateEventBridge();
				_AfterPresentationOpenEvent += value;
			}
			remove
			{
				_AfterPresentationOpenEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 11,12,14,15,16
		/// </summary>
		private event Application_PresentationSyncEventHandler _PresentationSyncEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744576.aspx </remarks>
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
		public event Application_PresentationSyncEventHandler PresentationSyncEvent
		{
			add
			{
				CreateEventBridge();
				_PresentationSyncEvent += value;
			}
			remove
			{
				_PresentationSyncEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 12,14,15,16
		/// </summary>
		private event Application_SlideShowOnNextEventHandler _SlideShowOnNextEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746469.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public event Application_SlideShowOnNextEventHandler SlideShowOnNextEvent
		{
			add
			{
				CreateEventBridge();
				_SlideShowOnNextEvent += value;
			}
			remove
			{
				_SlideShowOnNextEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 12,14,15,16
		/// </summary>
		private event Application_SlideShowOnPreviousEventHandler _SlideShowOnPreviousEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744749.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public event Application_SlideShowOnPreviousEventHandler SlideShowOnPreviousEvent
		{
			add
			{
				CreateEventBridge();
				_SlideShowOnPreviousEvent += value;
			}
			remove
			{
				_SlideShowOnPreviousEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 14,15,16
		/// </summary>
		private event Application_PresentationBeforeCloseEventHandler _PresentationBeforeCloseEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745567.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public event Application_PresentationBeforeCloseEventHandler PresentationBeforeCloseEvent
		{
			add
			{
				CreateEventBridge();
				_PresentationBeforeCloseEvent += value;
			}
			remove
			{
				_PresentationBeforeCloseEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 14,15,16
		/// </summary>
		private event Application_ProtectedViewWindowOpenEventHandler _ProtectedViewWindowOpenEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745081.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public event Application_ProtectedViewWindowOpenEventHandler ProtectedViewWindowOpenEvent
		{
			add
			{
				CreateEventBridge();
				_ProtectedViewWindowOpenEvent += value;
			}
			remove
			{
				_ProtectedViewWindowOpenEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 14,15,16
		/// </summary>
		private event Application_ProtectedViewWindowBeforeEditEventHandler _ProtectedViewWindowBeforeEditEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745575.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public event Application_ProtectedViewWindowBeforeEditEventHandler ProtectedViewWindowBeforeEditEvent
		{
			add
			{
				CreateEventBridge();
				_ProtectedViewWindowBeforeEditEvent += value;
			}
			remove
			{
				_ProtectedViewWindowBeforeEditEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 14,15,16
		/// </summary>
		private event Application_ProtectedViewWindowBeforeCloseEventHandler _ProtectedViewWindowBeforeCloseEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746497.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public event Application_ProtectedViewWindowBeforeCloseEventHandler ProtectedViewWindowBeforeCloseEvent
		{
			add
			{
				CreateEventBridge();
				_ProtectedViewWindowBeforeCloseEvent += value;
			}
			remove
			{
				_ProtectedViewWindowBeforeCloseEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 14,15,16
		/// </summary>
		private event Application_ProtectedViewWindowActivateEventHandler _ProtectedViewWindowActivateEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744591.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public event Application_ProtectedViewWindowActivateEventHandler ProtectedViewWindowActivateEvent
		{
			add
			{
				CreateEventBridge();
				_ProtectedViewWindowActivateEvent += value;
			}
			remove
			{
				_ProtectedViewWindowActivateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 14,15,16
		/// </summary>
		private event Application_ProtectedViewWindowDeactivateEventHandler _ProtectedViewWindowDeactivateEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746253.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public event Application_ProtectedViewWindowDeactivateEventHandler ProtectedViewWindowDeactivateEvent
		{
			add
			{
				CreateEventBridge();
				_ProtectedViewWindowDeactivateEvent += value;
			}
			remove
			{
				_ProtectedViewWindowDeactivateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 14,15,16
		/// </summary>
		private event Application_PresentationCloseFinalEventHandler _PresentationCloseFinalEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744781.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public event Application_PresentationCloseFinalEventHandler PresentationCloseFinalEvent
		{
			add
			{
				CreateEventBridge();
				_PresentationCloseFinalEvent += value;
			}
			remove
			{
				_PresentationCloseFinalEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 15, 16
		/// </summary>
		private event Application_AfterDragDropOnSlideEventHandler _AfterDragDropOnSlideEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj227644.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		public event Application_AfterDragDropOnSlideEventHandler AfterDragDropOnSlideEvent
		{
			add
			{
				CreateEventBridge();
				_AfterDragDropOnSlideEvent += value;
			}
			remove
			{
				_AfterDragDropOnSlideEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint, 15, 16
		/// </summary>
		private event Application_AfterShapeSizeChangeEventHandler _AfterShapeSizeChangeEvent;

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj227375.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		public event Application_AfterShapeSizeChangeEventHandler AfterShapeSizeChangeEvent
		{
			add
			{
				CreateEventBridge();
				_AfterShapeSizeChangeEvent += value;
			}
			remove
			{
				_AfterShapeSizeChangeEvent -= value;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, EApplication_SinkHelper.Id);


			if(EApplication_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_eApplication_SinkHelper = new EApplication_SinkHelper(this, _connectPoint);
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
			if( null != _eApplication_SinkHelper)
			{
				_eApplication_SinkHelper.Dispose();
				_eApplication_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}