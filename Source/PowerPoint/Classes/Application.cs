﻿using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.PowerPointApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Application_WindowSelectionChangeEventHandler(NetOffice.PowerPointApi.Selection sel);
	public delegate void Application_WindowBeforeRightClickEventHandler(NetOffice.PowerPointApi.Selection sel, ref bool cancel);
	public delegate void Application_WindowBeforeDoubleClickEventHandler(NetOffice.PowerPointApi.Selection sel, ref bool cancel);
	public delegate void Application_PresentationCloseEventHandler(NetOffice.PowerPointApi.Presentation pres);
	public delegate void Application_PresentationSaveEventHandler(NetOffice.PowerPointApi.Presentation pres);
	public delegate void Application_PresentationOpenEventHandler(NetOffice.PowerPointApi.Presentation pres);
	public delegate void Application_NewPresentationEventHandler(NetOffice.PowerPointApi.Presentation pres);
	public delegate void Application_PresentationNewSlideEventHandler(NetOffice.PowerPointApi.Slide sld);
	public delegate void Application_WindowActivateEventHandler(NetOffice.PowerPointApi.Presentation pres, NetOffice.PowerPointApi.DocumentWindow wn);
	public delegate void Application_WindowDeactivateEventHandler(NetOffice.PowerPointApi.Presentation pres, NetOffice.PowerPointApi.DocumentWindow wn);
	public delegate void Application_SlideShowBeginEventHandler(NetOffice.PowerPointApi.SlideShowWindow wn);
	public delegate void Application_SlideShowNextBuildEventHandler(NetOffice.PowerPointApi.SlideShowWindow wn);
	public delegate void Application_SlideShowNextSlideEventHandler(NetOffice.PowerPointApi.SlideShowWindow wn);
	public delegate void Application_SlideShowEndEventHandler(NetOffice.PowerPointApi.Presentation pres);
	public delegate void Application_PresentationPrintEventHandler(NetOffice.PowerPointApi.Presentation pres);
	public delegate void Application_SlideSelectionChangedEventHandler(NetOffice.PowerPointApi.SlideRange sldRange);
	public delegate void Application_ColorSchemeChangedEventHandler(NetOffice.PowerPointApi.SlideRange sldRange);
	public delegate void Application_PresentationBeforeSaveEventHandler(NetOffice.PowerPointApi.Presentation pres, ref bool Cancel);
	public delegate void Application_SlideShowNextClickEventHandler(NetOffice.PowerPointApi.SlideShowWindow wn, NetOffice.PowerPointApi.Effect nEffect);
	public delegate void Application_AfterNewPresentationEventHandler(NetOffice.PowerPointApi.Presentation pres);
	public delegate void Application_AfterPresentationOpenEventHandler(NetOffice.PowerPointApi.Presentation pres);
	public delegate void Application_PresentationSyncEventHandler(NetOffice.PowerPointApi.Presentation pres, NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType);
	public delegate void Application_SlideShowOnNextEventHandler(NetOffice.PowerPointApi.SlideShowWindow wn);
	public delegate void Application_SlideShowOnPreviousEventHandler(NetOffice.PowerPointApi.SlideShowWindow wn);
	public delegate void Application_PresentationBeforeCloseEventHandler(NetOffice.PowerPointApi.Presentation pres, ref bool cancel);
	public delegate void Application_ProtectedViewWindowOpenEventHandler(NetOffice.PowerPointApi.ProtectedViewWindow protViewWindow);
	public delegate void Application_ProtectedViewWindowBeforeEditEventHandler(NetOffice.PowerPointApi.ProtectedViewWindow protViewWindow, ref bool Cancel);
	public delegate void Application_ProtectedViewWindowBeforeCloseEventHandler(NetOffice.PowerPointApi.ProtectedViewWindow protViewWindow, NetOffice.PowerPointApi.Enums.PpProtectedViewCloseReason protectedViewCloseReason, ref bool cancel);
	public delegate void Application_ProtectedViewWindowActivateEventHandler(NetOffice.PowerPointApi.ProtectedViewWindow protViewWindow);
	public delegate void Application_ProtectedViewWindowDeactivateEventHandler(NetOffice.PowerPointApi.ProtectedViewWindow protViewWindow);
	public delegate void Application_PresentationCloseFinalEventHandler(NetOffice.PowerPointApi.Presentation pres);
	public delegate void Application_AfterDragDropOnSlideEventHandler(NetOffice.PowerPointApi.Slide sld, Single x, Single yY);
	public delegate void Application_AfterShapeSizeChangeEventHandler(NetOffice.PowerPointApi.Shape shp);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Application 
	/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application"/> </remarks>
	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass), ComProgId("PowerPoint.Application"), ModuleProvider(typeof(GlobalHelperModules.GlobalModule))]
	[EventSink(typeof(Events.EApplication_SinkHelper))]
    [ComEventInterface(typeof(Events.EApplication))]
    public class Application : _Application, ICloneable<Application>, IEventBinding
	{
		#pragma warning disable

		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;
        private Events.EApplication_SinkHelper _eApplication_SinkHelper;
	
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

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Application(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
			_callQuitInDispose = true;
		}

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
            _callQuitInDispose = true;
            GlobalHelperModules.GlobalModule.Instance = this;
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
                object proxy = Running.ProxyService.GetActiveInstance("PowerPoint", "Application", false);
                if (null != proxy)
                {
                    CreateFromProxy(proxy, true);
                    FromProxyService = true;
                }
                else
                {
                    CreateFromProgId("PowerPoint.Application", true);
                }
            }
            else
            {
                CreateFromProgId("PowerPoint.Application", true);
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
        /// Returns all running PowerPoint.Application instances from the environment/system
        /// </summary>
        /// <returns>PowerPoint.Application sequence</returns>
        public static IDisposableSequence<Application> GetActiveInstances()
        {
            return Running.ProxyService.GetActiveInstances<Application>("PowerPoint", "Application");
        }

        /// <summary>
        /// Returns a running PowerPoint.Application instance from the environment/system
        /// </summary>
        /// <param name="throwExceptionIfNotFound">throw exception if unable to find an instance</param>
        /// <returns>PowerPoint.Application instance or null</returns>
        public static Application GetActiveInstance(bool throwExceptionIfNotFound = false)
        {
            return Running.ProxyService.GetActiveInstance<Application>("PowerPoint", "Application", throwExceptionIfNotFound);
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.WindowSelectionChange"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.WindowBeforeRightClick"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.WindowBeforeDoubleClick"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.PresentationClose"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.PresentationSave"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.PresentationOpen"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.NewPresentation(even)"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.PresentationNewSlide"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.WindowActivate"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.WindowDeactivate"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.SlideShowBegin"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.SlideShowNextBuild"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.SlideShowNextSlide"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.SlideShowEnd"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.PresentationPrint"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.SlideSelectionChanged"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.ColorSchemeChanged"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.PresentationBeforeSave"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.SlideShowNextClick"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.AfterNewPresentation"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.AfterPresentationOpen"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.PresentationSync"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.SlideShowOnNext"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.SlideShowOnPrevious"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.PresentationBeforeClose"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.ProtectedViewWindowOpen"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.ProtectedViewWindowBeforeEdit"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.ProtectedViewWindowBeforeClose"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.ProtectedViewWindowActivate"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.ProtectedViewWindowDeactivate"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.PresentationCloseFinal"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.application.afterdragdroponslide"/> </remarks>
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
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.application.aftershapesizechange"/> </remarks>
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, Events.EApplication_SinkHelper.Id);


			if(Events.EApplication_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_eApplication_SinkHelper = new Events.EApplication_SinkHelper(this, _connectPoint);
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
			if( null != _eApplication_SinkHelper)
			{
				_eApplication_SinkHelper.Dispose();
				_eApplication_SinkHelper = null;
			}

			_connectPoint = null;
		}

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