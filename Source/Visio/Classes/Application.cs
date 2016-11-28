using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
using NetOffice.Misc;

namespace NetOffice.VisioApi
{

	#region Delegates

	#pragma warning disable
	public delegate void Application_AppActivatedEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_AppDeactivatedEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_AppObjActivatedEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_AppObjDeactivatedEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_BeforeQuitEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_BeforeModalEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_AfterModalEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_WindowOpenedEventHandler(NetOffice.VisioApi.IVWindow Window);
	public delegate void Application_SelectionChangedEventHandler(NetOffice.VisioApi.IVWindow Window);
	public delegate void Application_BeforeWindowClosedEventHandler(NetOffice.VisioApi.IVWindow Window);
	public delegate void Application_WindowActivatedEventHandler(NetOffice.VisioApi.IVWindow Window);
	public delegate void Application_BeforeWindowSelDeleteEventHandler(NetOffice.VisioApi.IVWindow Window);
	public delegate void Application_BeforeWindowPageTurnEventHandler(NetOffice.VisioApi.IVWindow Window);
	public delegate void Application_WindowTurnedToPageEventHandler(NetOffice.VisioApi.IVWindow Window);
	public delegate void Application_DocumentOpenedEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Application_DocumentCreatedEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Application_DocumentSavedEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Application_DocumentSavedAsEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Application_DocumentChangedEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Application_BeforeDocumentCloseEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Application_StyleAddedEventHandler(NetOffice.VisioApi.IVStyle Style);
	public delegate void Application_StyleChangedEventHandler(NetOffice.VisioApi.IVStyle Style);
	public delegate void Application_BeforeStyleDeleteEventHandler(NetOffice.VisioApi.IVStyle Style);
	public delegate void Application_MasterAddedEventHandler(NetOffice.VisioApi.IVMaster Master);
	public delegate void Application_MasterChangedEventHandler(NetOffice.VisioApi.IVMaster Master);
	public delegate void Application_BeforeMasterDeleteEventHandler(NetOffice.VisioApi.IVMaster Master);
	public delegate void Application_PageAddedEventHandler(NetOffice.VisioApi.IVPage Page);
	public delegate void Application_PageChangedEventHandler(NetOffice.VisioApi.IVPage Page);
	public delegate void Application_BeforePageDeleteEventHandler(NetOffice.VisioApi.IVPage Page);
	public delegate void Application_ShapeAddedEventHandler(NetOffice.VisioApi.IVShape Shape);
	public delegate void Application_BeforeSelectionDeleteEventHandler(NetOffice.VisioApi.IVSelection Selection);
	public delegate void Application_ShapeChangedEventHandler(NetOffice.VisioApi.IVShape Shape);
	public delegate void Application_SelectionAddedEventHandler(NetOffice.VisioApi.IVSelection Selection);
	public delegate void Application_BeforeShapeDeleteEventHandler(NetOffice.VisioApi.IVShape Shape);
	public delegate void Application_TextChangedEventHandler(NetOffice.VisioApi.IVShape Shape);
	public delegate void Application_CellChangedEventHandler(NetOffice.VisioApi.IVCell Cell);
	public delegate void Application_MarkerEventEventHandler(NetOffice.VisioApi.IVApplication app, Int32 SequenceNum, string ContextString);
	public delegate void Application_NoEventsPendingEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_VisioIsIdleEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_MustFlushScopeBeginningEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_MustFlushScopeEndedEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_RunModeEnteredEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Application_DesignModeEnteredEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Application_BeforeDocumentSaveEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Application_BeforeDocumentSaveAsEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Application_FormulaChangedEventHandler(NetOffice.VisioApi.IVCell Cell);
	public delegate void Application_ConnectionsAddedEventHandler(NetOffice.VisioApi.IVConnects Connects);
	public delegate void Application_ConnectionsDeletedEventHandler(NetOffice.VisioApi.IVConnects Connects);
	public delegate void Application_EnterScopeEventHandler(NetOffice.VisioApi.IVApplication app, Int32 nScopeID, string bstrDescription);
	public delegate void Application_ExitScopeEventHandler(NetOffice.VisioApi.IVApplication app, Int32 nScopeID, string bstrDescription, bool bErrOrCancelled);
	public delegate void Application_QueryCancelQuitEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_QuitCanceledEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_WindowChangedEventHandler(NetOffice.VisioApi.IVWindow Window);
	public delegate void Application_ViewChangedEventHandler(NetOffice.VisioApi.IVWindow Window);
	public delegate void Application_QueryCancelWindowCloseEventHandler(NetOffice.VisioApi.IVWindow Window);
	public delegate void Application_WindowCloseCanceledEventHandler(NetOffice.VisioApi.IVWindow Window);
	public delegate void Application_QueryCancelDocumentCloseEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Application_DocumentCloseCanceledEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Application_QueryCancelStyleDeleteEventHandler(NetOffice.VisioApi.IVStyle Style);
	public delegate void Application_StyleDeleteCanceledEventHandler(NetOffice.VisioApi.IVStyle Style);
	public delegate void Application_QueryCancelMasterDeleteEventHandler(NetOffice.VisioApi.IVMaster Master);
	public delegate void Application_MasterDeleteCanceledEventHandler(NetOffice.VisioApi.IVMaster Master);
	public delegate void Application_QueryCancelPageDeleteEventHandler(NetOffice.VisioApi.IVPage Page);
	public delegate void Application_PageDeleteCanceledEventHandler(NetOffice.VisioApi.IVPage Page);
	public delegate void Application_ShapeParentChangedEventHandler(NetOffice.VisioApi.IVShape Shape);
	public delegate void Application_BeforeShapeTextEditEventHandler(NetOffice.VisioApi.IVShape Shape);
	public delegate void Application_ShapeExitedTextEditEventHandler(NetOffice.VisioApi.IVShape Shape);
	public delegate void Application_QueryCancelSelectionDeleteEventHandler(NetOffice.VisioApi.IVSelection Selection);
	public delegate void Application_SelectionDeleteCanceledEventHandler(NetOffice.VisioApi.IVSelection Selection);
	public delegate void Application_QueryCancelUngroupEventHandler(NetOffice.VisioApi.IVSelection Selection);
	public delegate void Application_UngroupCanceledEventHandler(NetOffice.VisioApi.IVSelection Selection);
	public delegate void Application_QueryCancelConvertToGroupEventHandler(NetOffice.VisioApi.IVSelection Selection);
	public delegate void Application_ConvertToGroupCanceledEventHandler(NetOffice.VisioApi.IVSelection Selection);
	public delegate void Application_QueryCancelSuspendEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_SuspendCanceledEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_BeforeSuspendEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_AfterResumeEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_OnKeystrokeMessageForAddonEventHandler(NetOffice.VisioApi.IVMSGWrap MSG);
	public delegate void Application_MouseDownEventHandler(Int32 Button, Int32 KeyButtonState, Double x, Double y, ref bool CancelDefault);
	public delegate void Application_MouseMoveEventHandler(Int32 Button, Int32 KeyButtonState, Double x, Double y, ref bool CancelDefault);
	public delegate void Application_MouseUpEventHandler(Int32 Button, Int32 KeyButtonState, Double x, Double y, ref bool CancelDefault);
	public delegate void Application_KeyDownEventHandler(Int32 KeyCode, Int32 KeyButtonState, ref bool CancelDefault);
	public delegate void Application_KeyPressEventHandler(Int32 KeyAscii, ref bool CancelDefault);
	public delegate void Application_KeyUpEventHandler(Int32 KeyCode, Int32 KeyButtonState, ref bool CancelDefault);
	public delegate void Application_QueryCancelSuspendEventsEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_SuspendEventsCanceledEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_BeforeSuspendEventsEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_AfterResumeEventsEventHandler(NetOffice.VisioApi.IVApplication app);
	public delegate void Application_QueryCancelGroupEventHandler(NetOffice.VisioApi.IVSelection Selection);
	public delegate void Application_GroupCanceledEventHandler(NetOffice.VisioApi.IVSelection Selection);
	public delegate void Application_ShapeDataGraphicChangedEventHandler(NetOffice.VisioApi.IVShape Shape);
	public delegate void Application_BeforeDataRecordsetDeleteEventHandler(NetOffice.VisioApi.IVDataRecordset DataRecordset);
	public delegate void Application_DataRecordsetChangedEventHandler(NetOffice.VisioApi.IVDataRecordsetChangedEvent DataRecordsetChanged);
	public delegate void Application_DataRecordsetAddedEventHandler(NetOffice.VisioApi.IVDataRecordset DataRecordset);
	public delegate void Application_ShapeLinkAddedEventHandler(NetOffice.VisioApi.IVShape Shape, Int32 DataRecordsetID, Int32 DataRowID);
	public delegate void Application_ShapeLinkDeletedEventHandler(NetOffice.VisioApi.IVShape Shape, Int32 DataRecordsetID, Int32 DataRowID);
	public delegate void Application_AfterRemoveHiddenInformationEventHandler(NetOffice.VisioApi.IVDocument doc);
	public delegate void Application_ContainerRelationshipAddedEventHandler(NetOffice.VisioApi.IVRelatedShapePairEvent ShapePair);
	public delegate void Application_ContainerRelationshipDeletedEventHandler(NetOffice.VisioApi.IVRelatedShapePairEvent ShapePair);
	public delegate void Application_CalloutRelationshipAddedEventHandler(NetOffice.VisioApi.IVRelatedShapePairEvent ShapePair);
	public delegate void Application_CalloutRelationshipDeletedEventHandler(NetOffice.VisioApi.IVRelatedShapePairEvent ShapePair);
	public delegate void Application_RuleSetValidatedEventHandler(NetOffice.VisioApi.IVValidationRuleSet RuleSet);
	public delegate void Application_QueryCancelReplaceShapesEventHandler(NetOffice.VisioApi.IVReplaceShapesEvent replaceShapes);
	public delegate void Application_ReplaceShapesCanceledEventHandler(NetOffice.VisioApi.IVReplaceShapesEvent replaceShapes);
	public delegate void Application_BeforeReplaceShapesEventHandler(NetOffice.VisioApi.IVReplaceShapesEvent replaceShapes);
	public delegate void Application_AfterReplaceShapesEventHandler(NetOffice.VisioApi.IVSelection sel);
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass Application 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769220(v=office.14).aspx
	///</summary>
	[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class Application : IVApplication,IEventBinding
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
			
			GlobalHelperModules.GlobalModule.Instance = this;
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Application(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
			GlobalHelperModules.GlobalModule.Instance = this;
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Application(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Application(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Application(ICOMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		///<summary>
        /// Creates a new instance of Application 
        ///</summary>		
		public Application():base("Visio.Application")
		{
			
			GlobalHelperModules.GlobalModule.Instance = this;
		}
		
		///<summary>
        /// Creates a new instance of Application
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public Application(string progId):base(progId)
		{
			
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
        /// Returns all running Visio.Application objects from the environment/system
        /// </summary>
        /// <returns>an Visio.Application array</returns>
		public static NetOffice.VisioApi.Application[] GetActiveInstances()
		{		
			IDisposableEnumeration proxyList = NetOffice.ProxyService.GetActiveInstances("Visio","Application");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.VisioApi.Application> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.VisioApi.Application>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.VisioApi.Application(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// Returns a running Visio.Application object from the environment/system.
        /// </summary>
        /// <returns>an Visio.Application object or null</returns>
		public static NetOffice.VisioApi.Application GetActiveInstance()
		{
			object proxy  = NetOffice.ProxyService.GetActiveInstance("Visio","Application", false);
			if(null != proxy)
				return new NetOffice.VisioApi.Application(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// Returns a running Visio.Application object from the environment/system. 
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an Visio.Application object or null</returns>
		public static NetOffice.VisioApi.Application GetActiveInstance(bool throwOnError)
		{
			object proxy  = NetOffice.ProxyService.GetActiveInstance("Visio","Application", throwOnError);
			if(null != proxy)
				return new NetOffice.VisioApi.Application(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_AppActivatedEventHandler _AppActivatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765356(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_AppActivatedEventHandler AppActivatedEvent
		{
			add
			{
				CreateEventBridge();
				_AppActivatedEvent += value;
			}
			remove
			{
				_AppActivatedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_AppDeactivatedEventHandler _AppDeactivatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765903(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_AppDeactivatedEventHandler AppDeactivatedEvent
		{
			add
			{
				CreateEventBridge();
				_AppDeactivatedEvent += value;
			}
			remove
			{
				_AppDeactivatedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_AppObjActivatedEventHandler _AppObjActivatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767797(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_AppObjActivatedEventHandler AppObjActivatedEvent
		{
			add
			{
				CreateEventBridge();
				_AppObjActivatedEvent += value;
			}
			remove
			{
				_AppObjActivatedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_AppObjDeactivatedEventHandler _AppObjDeactivatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765196(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_AppObjDeactivatedEventHandler AppObjDeactivatedEvent
		{
			add
			{
				CreateEventBridge();
				_AppObjDeactivatedEvent += value;
			}
			remove
			{
				_AppObjDeactivatedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_BeforeQuitEventHandler _BeforeQuitEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767832(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_BeforeQuitEventHandler BeforeQuitEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeQuitEvent += value;
			}
			remove
			{
				_BeforeQuitEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_BeforeModalEventHandler _BeforeModalEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766316(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_BeforeModalEventHandler BeforeModalEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeModalEvent += value;
			}
			remove
			{
				_BeforeModalEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_AfterModalEventHandler _AfterModalEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768670(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_AfterModalEventHandler AfterModalEvent
		{
			add
			{
				CreateEventBridge();
				_AfterModalEvent += value;
			}
			remove
			{
				_AfterModalEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_WindowOpenedEventHandler _WindowOpenedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767725(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_WindowOpenedEventHandler WindowOpenedEvent
		{
			add
			{
				CreateEventBridge();
				_WindowOpenedEvent += value;
			}
			remove
			{
				_WindowOpenedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_SelectionChangedEventHandler _SelectionChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768425(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_SelectionChangedEventHandler SelectionChangedEvent
		{
			add
			{
				CreateEventBridge();
				_SelectionChangedEvent += value;
			}
			remove
			{
				_SelectionChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_BeforeWindowClosedEventHandler _BeforeWindowClosedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768648(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_BeforeWindowClosedEventHandler BeforeWindowClosedEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeWindowClosedEvent += value;
			}
			remove
			{
				_BeforeWindowClosedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_WindowActivatedEventHandler _WindowActivatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768932(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_WindowActivatedEventHandler WindowActivatedEvent
		{
			add
			{
				CreateEventBridge();
				_WindowActivatedEvent += value;
			}
			remove
			{
				_WindowActivatedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_BeforeWindowSelDeleteEventHandler _BeforeWindowSelDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765921(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_BeforeWindowSelDeleteEventHandler BeforeWindowSelDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeWindowSelDeleteEvent += value;
			}
			remove
			{
				_BeforeWindowSelDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_BeforeWindowPageTurnEventHandler _BeforeWindowPageTurnEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768604(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_BeforeWindowPageTurnEventHandler BeforeWindowPageTurnEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeWindowPageTurnEvent += value;
			}
			remove
			{
				_BeforeWindowPageTurnEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_WindowTurnedToPageEventHandler _WindowTurnedToPageEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769065(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_WindowTurnedToPageEventHandler WindowTurnedToPageEvent
		{
			add
			{
				CreateEventBridge();
				_WindowTurnedToPageEvent += value;
			}
			remove
			{
				_WindowTurnedToPageEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_DocumentOpenedEventHandler _DocumentOpenedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768552(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_DocumentOpenedEventHandler DocumentOpenedEvent
		{
			add
			{
				CreateEventBridge();
				_DocumentOpenedEvent += value;
			}
			remove
			{
				_DocumentOpenedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_DocumentCreatedEventHandler _DocumentCreatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765843(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_DocumentCreatedEventHandler DocumentCreatedEvent
		{
			add
			{
				CreateEventBridge();
				_DocumentCreatedEvent += value;
			}
			remove
			{
				_DocumentCreatedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_DocumentSavedEventHandler _DocumentSavedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767627(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_DocumentSavedEventHandler DocumentSavedEvent
		{
			add
			{
				CreateEventBridge();
				_DocumentSavedEvent += value;
			}
			remove
			{
				_DocumentSavedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_DocumentSavedAsEventHandler _DocumentSavedAsEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768947(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_DocumentSavedAsEventHandler DocumentSavedAsEvent
		{
			add
			{
				CreateEventBridge();
				_DocumentSavedAsEvent += value;
			}
			remove
			{
				_DocumentSavedAsEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_DocumentChangedEventHandler _DocumentChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768119(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_DocumentChangedEventHandler DocumentChangedEvent
		{
			add
			{
				CreateEventBridge();
				_DocumentChangedEvent += value;
			}
			remove
			{
				_DocumentChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_BeforeDocumentCloseEventHandler _BeforeDocumentCloseEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768151(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_BeforeDocumentCloseEventHandler BeforeDocumentCloseEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeDocumentCloseEvent += value;
			}
			remove
			{
				_BeforeDocumentCloseEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_StyleAddedEventHandler _StyleAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767751(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_StyleAddedEventHandler StyleAddedEvent
		{
			add
			{
				CreateEventBridge();
				_StyleAddedEvent += value;
			}
			remove
			{
				_StyleAddedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_StyleChangedEventHandler _StyleChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769029(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_StyleChangedEventHandler StyleChangedEvent
		{
			add
			{
				CreateEventBridge();
				_StyleChangedEvent += value;
			}
			remove
			{
				_StyleChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_BeforeStyleDeleteEventHandler _BeforeStyleDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766539(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_BeforeStyleDeleteEventHandler BeforeStyleDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeStyleDeleteEvent += value;
			}
			remove
			{
				_BeforeStyleDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_MasterAddedEventHandler _MasterAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768930(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_MasterAddedEventHandler MasterAddedEvent
		{
			add
			{
				CreateEventBridge();
				_MasterAddedEvent += value;
			}
			remove
			{
				_MasterAddedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_MasterChangedEventHandler _MasterChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769090(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_MasterChangedEventHandler MasterChangedEvent
		{
			add
			{
				CreateEventBridge();
				_MasterChangedEvent += value;
			}
			remove
			{
				_MasterChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_BeforeMasterDeleteEventHandler _BeforeMasterDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766726(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_BeforeMasterDeleteEventHandler BeforeMasterDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeMasterDeleteEvent += value;
			}
			remove
			{
				_BeforeMasterDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_PageAddedEventHandler _PageAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765378(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_PageAddedEventHandler PageAddedEvent
		{
			add
			{
				CreateEventBridge();
				_PageAddedEvent += value;
			}
			remove
			{
				_PageAddedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_PageChangedEventHandler _PageChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768083(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_PageChangedEventHandler PageChangedEvent
		{
			add
			{
				CreateEventBridge();
				_PageChangedEvent += value;
			}
			remove
			{
				_PageChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_BeforePageDeleteEventHandler _BeforePageDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766722(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_BeforePageDeleteEventHandler BeforePageDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_BeforePageDeleteEvent += value;
			}
			remove
			{
				_BeforePageDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_ShapeAddedEventHandler _ShapeAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766392(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_ShapeAddedEventHandler ShapeAddedEvent
		{
			add
			{
				CreateEventBridge();
				_ShapeAddedEvent += value;
			}
			remove
			{
				_ShapeAddedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_BeforeSelectionDeleteEventHandler _BeforeSelectionDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766131(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_BeforeSelectionDeleteEventHandler BeforeSelectionDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeSelectionDeleteEvent += value;
			}
			remove
			{
				_BeforeSelectionDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_ShapeChangedEventHandler _ShapeChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767789(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_ShapeChangedEventHandler ShapeChangedEvent
		{
			add
			{
				CreateEventBridge();
				_ShapeChangedEvent += value;
			}
			remove
			{
				_ShapeChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_SelectionAddedEventHandler _SelectionAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766974(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_SelectionAddedEventHandler SelectionAddedEvent
		{
			add
			{
				CreateEventBridge();
				_SelectionAddedEvent += value;
			}
			remove
			{
				_SelectionAddedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_BeforeShapeDeleteEventHandler _BeforeShapeDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767938(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_BeforeShapeDeleteEventHandler BeforeShapeDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeShapeDeleteEvent += value;
			}
			remove
			{
				_BeforeShapeDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_TextChangedEventHandler _TextChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767908(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_TextChangedEventHandler TextChangedEvent
		{
			add
			{
				CreateEventBridge();
				_TextChangedEvent += value;
			}
			remove
			{
				_TextChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_CellChangedEventHandler _CellChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767326(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_CellChangedEventHandler CellChangedEvent
		{
			add
			{
				CreateEventBridge();
				_CellChangedEvent += value;
			}
			remove
			{
				_CellChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_MarkerEventEventHandler _MarkerEventEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765486(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_MarkerEventEventHandler MarkerEventEvent
		{
			add
			{
				CreateEventBridge();
				_MarkerEventEvent += value;
			}
			remove
			{
				_MarkerEventEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_NoEventsPendingEventHandler _NoEventsPendingEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767335(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_NoEventsPendingEventHandler NoEventsPendingEvent
		{
			add
			{
				CreateEventBridge();
				_NoEventsPendingEvent += value;
			}
			remove
			{
				_NoEventsPendingEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_VisioIsIdleEventHandler _VisioIsIdleEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766439(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_VisioIsIdleEventHandler VisioIsIdleEvent
		{
			add
			{
				CreateEventBridge();
				_VisioIsIdleEvent += value;
			}
			remove
			{
				_VisioIsIdleEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_MustFlushScopeBeginningEventHandler _MustFlushScopeBeginningEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767512(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_MustFlushScopeBeginningEventHandler MustFlushScopeBeginningEvent
		{
			add
			{
				CreateEventBridge();
				_MustFlushScopeBeginningEvent += value;
			}
			remove
			{
				_MustFlushScopeBeginningEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_MustFlushScopeEndedEventHandler _MustFlushScopeEndedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768052(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_MustFlushScopeEndedEventHandler MustFlushScopeEndedEvent
		{
			add
			{
				CreateEventBridge();
				_MustFlushScopeEndedEvent += value;
			}
			remove
			{
				_MustFlushScopeEndedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_RunModeEnteredEventHandler _RunModeEnteredEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765975(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_RunModeEnteredEventHandler RunModeEnteredEvent
		{
			add
			{
				CreateEventBridge();
				_RunModeEnteredEvent += value;
			}
			remove
			{
				_RunModeEnteredEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_DesignModeEnteredEventHandler _DesignModeEnteredEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765830(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_DesignModeEnteredEventHandler DesignModeEnteredEvent
		{
			add
			{
				CreateEventBridge();
				_DesignModeEnteredEvent += value;
			}
			remove
			{
				_DesignModeEnteredEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_BeforeDocumentSaveEventHandler _BeforeDocumentSaveEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768487(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_BeforeDocumentSaveEventHandler BeforeDocumentSaveEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeDocumentSaveEvent += value;
			}
			remove
			{
				_BeforeDocumentSaveEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_BeforeDocumentSaveAsEventHandler _BeforeDocumentSaveAsEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768752(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_BeforeDocumentSaveAsEventHandler BeforeDocumentSaveAsEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeDocumentSaveAsEvent += value;
			}
			remove
			{
				_BeforeDocumentSaveAsEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_FormulaChangedEventHandler _FormulaChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769046(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_FormulaChangedEventHandler FormulaChangedEvent
		{
			add
			{
				CreateEventBridge();
				_FormulaChangedEvent += value;
			}
			remove
			{
				_FormulaChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_ConnectionsAddedEventHandler _ConnectionsAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768100(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_ConnectionsAddedEventHandler ConnectionsAddedEvent
		{
			add
			{
				CreateEventBridge();
				_ConnectionsAddedEvent += value;
			}
			remove
			{
				_ConnectionsAddedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_ConnectionsDeletedEventHandler _ConnectionsDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767479(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_ConnectionsDeletedEventHandler ConnectionsDeletedEvent
		{
			add
			{
				CreateEventBridge();
				_ConnectionsDeletedEvent += value;
			}
			remove
			{
				_ConnectionsDeletedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_EnterScopeEventHandler _EnterScopeEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769070(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_EnterScopeEventHandler EnterScopeEvent
		{
			add
			{
				CreateEventBridge();
				_EnterScopeEvent += value;
			}
			remove
			{
				_EnterScopeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_ExitScopeEventHandler _ExitScopeEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767448(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_ExitScopeEventHandler ExitScopeEvent
		{
			add
			{
				CreateEventBridge();
				_ExitScopeEvent += value;
			}
			remove
			{
				_ExitScopeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_QueryCancelQuitEventHandler _QueryCancelQuitEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765429(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_QueryCancelQuitEventHandler QueryCancelQuitEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelQuitEvent += value;
			}
			remove
			{
				_QueryCancelQuitEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_QuitCanceledEventHandler _QuitCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765158(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_QuitCanceledEventHandler QuitCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_QuitCanceledEvent += value;
			}
			remove
			{
				_QuitCanceledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_WindowChangedEventHandler _WindowChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765706(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_WindowChangedEventHandler WindowChangedEvent
		{
			add
			{
				CreateEventBridge();
				_WindowChangedEvent += value;
			}
			remove
			{
				_WindowChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_ViewChangedEventHandler _ViewChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765751(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_ViewChangedEventHandler ViewChangedEvent
		{
			add
			{
				CreateEventBridge();
				_ViewChangedEvent += value;
			}
			remove
			{
				_ViewChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_QueryCancelWindowCloseEventHandler _QueryCancelWindowCloseEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769020(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_QueryCancelWindowCloseEventHandler QueryCancelWindowCloseEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelWindowCloseEvent += value;
			}
			remove
			{
				_QueryCancelWindowCloseEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_WindowCloseCanceledEventHandler _WindowCloseCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765316(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_WindowCloseCanceledEventHandler WindowCloseCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_WindowCloseCanceledEvent += value;
			}
			remove
			{
				_WindowCloseCanceledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_QueryCancelDocumentCloseEventHandler _QueryCancelDocumentCloseEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766512(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_QueryCancelDocumentCloseEventHandler QueryCancelDocumentCloseEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelDocumentCloseEvent += value;
			}
			remove
			{
				_QueryCancelDocumentCloseEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_DocumentCloseCanceledEventHandler _DocumentCloseCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765332(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_DocumentCloseCanceledEventHandler DocumentCloseCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_DocumentCloseCanceledEvent += value;
			}
			remove
			{
				_DocumentCloseCanceledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_QueryCancelStyleDeleteEventHandler _QueryCancelStyleDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767117(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_QueryCancelStyleDeleteEventHandler QueryCancelStyleDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelStyleDeleteEvent += value;
			}
			remove
			{
				_QueryCancelStyleDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_StyleDeleteCanceledEventHandler _StyleDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768236(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_StyleDeleteCanceledEventHandler StyleDeleteCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_StyleDeleteCanceledEvent += value;
			}
			remove
			{
				_StyleDeleteCanceledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_QueryCancelMasterDeleteEventHandler _QueryCancelMasterDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767170(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_QueryCancelMasterDeleteEventHandler QueryCancelMasterDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelMasterDeleteEvent += value;
			}
			remove
			{
				_QueryCancelMasterDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_MasterDeleteCanceledEventHandler _MasterDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767357(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_MasterDeleteCanceledEventHandler MasterDeleteCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_MasterDeleteCanceledEvent += value;
			}
			remove
			{
				_MasterDeleteCanceledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_QueryCancelPageDeleteEventHandler _QueryCancelPageDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767161(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_QueryCancelPageDeleteEventHandler QueryCancelPageDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelPageDeleteEvent += value;
			}
			remove
			{
				_QueryCancelPageDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_PageDeleteCanceledEventHandler _PageDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765525(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_PageDeleteCanceledEventHandler PageDeleteCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_PageDeleteCanceledEvent += value;
			}
			remove
			{
				_PageDeleteCanceledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_ShapeParentChangedEventHandler _ShapeParentChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765842(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_ShapeParentChangedEventHandler ShapeParentChangedEvent
		{
			add
			{
				CreateEventBridge();
				_ShapeParentChangedEvent += value;
			}
			remove
			{
				_ShapeParentChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_BeforeShapeTextEditEventHandler _BeforeShapeTextEditEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768563(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_BeforeShapeTextEditEventHandler BeforeShapeTextEditEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeShapeTextEditEvent += value;
			}
			remove
			{
				_BeforeShapeTextEditEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_ShapeExitedTextEditEventHandler _ShapeExitedTextEditEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767733(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_ShapeExitedTextEditEventHandler ShapeExitedTextEditEvent
		{
			add
			{
				CreateEventBridge();
				_ShapeExitedTextEditEvent += value;
			}
			remove
			{
				_ShapeExitedTextEditEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_QueryCancelSelectionDeleteEventHandler _QueryCancelSelectionDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768575(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_QueryCancelSelectionDeleteEventHandler QueryCancelSelectionDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelSelectionDeleteEvent += value;
			}
			remove
			{
				_QueryCancelSelectionDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_SelectionDeleteCanceledEventHandler _SelectionDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766557(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_SelectionDeleteCanceledEventHandler SelectionDeleteCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_SelectionDeleteCanceledEvent += value;
			}
			remove
			{
				_SelectionDeleteCanceledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_QueryCancelUngroupEventHandler _QueryCancelUngroupEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766751(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_QueryCancelUngroupEventHandler QueryCancelUngroupEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelUngroupEvent += value;
			}
			remove
			{
				_QueryCancelUngroupEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_UngroupCanceledEventHandler _UngroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765728(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_UngroupCanceledEventHandler UngroupCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_UngroupCanceledEvent += value;
			}
			remove
			{
				_UngroupCanceledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_QueryCancelConvertToGroupEventHandler _QueryCancelConvertToGroupEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765456(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_QueryCancelConvertToGroupEventHandler QueryCancelConvertToGroupEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelConvertToGroupEvent += value;
			}
			remove
			{
				_QueryCancelConvertToGroupEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_ConvertToGroupCanceledEventHandler _ConvertToGroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765231(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_ConvertToGroupCanceledEventHandler ConvertToGroupCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_ConvertToGroupCanceledEvent += value;
			}
			remove
			{
				_ConvertToGroupCanceledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_QueryCancelSuspendEventHandler _QueryCancelSuspendEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765467(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_QueryCancelSuspendEventHandler QueryCancelSuspendEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelSuspendEvent += value;
			}
			remove
			{
				_QueryCancelSuspendEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_SuspendCanceledEventHandler _SuspendCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766700(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_SuspendCanceledEventHandler SuspendCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_SuspendCanceledEvent += value;
			}
			remove
			{
				_SuspendCanceledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_BeforeSuspendEventHandler _BeforeSuspendEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766733(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_BeforeSuspendEventHandler BeforeSuspendEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeSuspendEvent += value;
			}
			remove
			{
				_BeforeSuspendEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_AfterResumeEventHandler _AfterResumeEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766935(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_AfterResumeEventHandler AfterResumeEvent
		{
			add
			{
				CreateEventBridge();
				_AfterResumeEvent += value;
			}
			remove
			{
				_AfterResumeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_OnKeystrokeMessageForAddonEventHandler _OnKeystrokeMessageForAddonEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765211(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_OnKeystrokeMessageForAddonEventHandler OnKeystrokeMessageForAddonEvent
		{
			add
			{
				CreateEventBridge();
				_OnKeystrokeMessageForAddonEvent += value;
			}
			remove
			{
				_OnKeystrokeMessageForAddonEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_MouseDownEventHandler _MouseDownEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769048(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_MouseDownEventHandler MouseDownEvent
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
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_MouseMoveEventHandler _MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766075(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_MouseMoveEventHandler MouseMoveEvent
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
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_MouseUpEventHandler _MouseUpEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767334(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_MouseUpEventHandler MouseUpEvent
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
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_KeyDownEventHandler _KeyDownEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766050(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_KeyDownEventHandler KeyDownEvent
		{
			add
			{
				CreateEventBridge();
				_KeyDownEvent += value;
			}
			remove
			{
				_KeyDownEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_KeyPressEventHandler _KeyPressEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768385(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_KeyPressEventHandler KeyPressEvent
		{
			add
			{
				CreateEventBridge();
				_KeyPressEvent += value;
			}
			remove
			{
				_KeyPressEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Application_KeyUpEventHandler _KeyUpEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769131(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public event Application_KeyUpEventHandler KeyUpEvent
		{
			add
			{
				CreateEventBridge();
				_KeyUpEvent += value;
			}
			remove
			{
				_KeyUpEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 12,14,15,16
		/// </summary>
		private event Application_QueryCancelSuspendEventsEventHandler _QueryCancelSuspendEventsEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767255(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public event Application_QueryCancelSuspendEventsEventHandler QueryCancelSuspendEventsEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelSuspendEventsEvent += value;
			}
			remove
			{
				_QueryCancelSuspendEventsEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 12,14,15,16
		/// </summary>
		private event Application_SuspendEventsCanceledEventHandler _SuspendEventsCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765864(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public event Application_SuspendEventsCanceledEventHandler SuspendEventsCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_SuspendEventsCanceledEvent += value;
			}
			remove
			{
				_SuspendEventsCanceledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 12,14,15,16
		/// </summary>
		private event Application_BeforeSuspendEventsEventHandler _BeforeSuspendEventsEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767714(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public event Application_BeforeSuspendEventsEventHandler BeforeSuspendEventsEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeSuspendEventsEvent += value;
			}
			remove
			{
				_BeforeSuspendEventsEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 12,14,15,16
		/// </summary>
		private event Application_AfterResumeEventsEventHandler _AfterResumeEventsEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768214(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public event Application_AfterResumeEventsEventHandler AfterResumeEventsEvent
		{
			add
			{
				CreateEventBridge();
				_AfterResumeEventsEvent += value;
			}
			remove
			{
				_AfterResumeEventsEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 12,14,15,16
		/// </summary>
		private event Application_QueryCancelGroupEventHandler _QueryCancelGroupEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767917(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public event Application_QueryCancelGroupEventHandler QueryCancelGroupEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelGroupEvent += value;
			}
			remove
			{
				_QueryCancelGroupEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 12,14,15,16
		/// </summary>
		private event Application_GroupCanceledEventHandler _GroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768118(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public event Application_GroupCanceledEventHandler GroupCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_GroupCanceledEvent += value;
			}
			remove
			{
				_GroupCanceledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 12,14,15,16
		/// </summary>
		private event Application_ShapeDataGraphicChangedEventHandler _ShapeDataGraphicChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765725(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public event Application_ShapeDataGraphicChangedEventHandler ShapeDataGraphicChangedEvent
		{
			add
			{
				CreateEventBridge();
				_ShapeDataGraphicChangedEvent += value;
			}
			remove
			{
				_ShapeDataGraphicChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 12,14,15,16
		/// </summary>
		private event Application_BeforeDataRecordsetDeleteEventHandler _BeforeDataRecordsetDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767894(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public event Application_BeforeDataRecordsetDeleteEventHandler BeforeDataRecordsetDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeDataRecordsetDeleteEvent += value;
			}
			remove
			{
				_BeforeDataRecordsetDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 12,14,15,16
		/// </summary>
		private event Application_DataRecordsetChangedEventHandler _DataRecordsetChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767322(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public event Application_DataRecordsetChangedEventHandler DataRecordsetChangedEvent
		{
			add
			{
				CreateEventBridge();
				_DataRecordsetChangedEvent += value;
			}
			remove
			{
				_DataRecordsetChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 12,14,15,16
		/// </summary>
		private event Application_DataRecordsetAddedEventHandler _DataRecordsetAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765105(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public event Application_DataRecordsetAddedEventHandler DataRecordsetAddedEvent
		{
			add
			{
				CreateEventBridge();
				_DataRecordsetAddedEvent += value;
			}
			remove
			{
				_DataRecordsetAddedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 12,14,15,16
		/// </summary>
		private event Application_ShapeLinkAddedEventHandler _ShapeLinkAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765626(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public event Application_ShapeLinkAddedEventHandler ShapeLinkAddedEvent
		{
			add
			{
				CreateEventBridge();
				_ShapeLinkAddedEvent += value;
			}
			remove
			{
				_ShapeLinkAddedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 12,14,15,16
		/// </summary>
		private event Application_ShapeLinkDeletedEventHandler _ShapeLinkDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768165(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public event Application_ShapeLinkDeletedEventHandler ShapeLinkDeletedEvent
		{
			add
			{
				CreateEventBridge();
				_ShapeLinkDeletedEvent += value;
			}
			remove
			{
				_ShapeLinkDeletedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 12,14,15,16
		/// </summary>
		private event Application_AfterRemoveHiddenInformationEventHandler _AfterRemoveHiddenInformationEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767810(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public event Application_AfterRemoveHiddenInformationEventHandler AfterRemoveHiddenInformationEvent
		{
			add
			{
				CreateEventBridge();
				_AfterRemoveHiddenInformationEvent += value;
			}
			remove
			{
				_AfterRemoveHiddenInformationEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 14,15,16
		/// </summary>
		private event Application_ContainerRelationshipAddedEventHandler _ContainerRelationshipAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767352(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		public event Application_ContainerRelationshipAddedEventHandler ContainerRelationshipAddedEvent
		{
			add
			{
				CreateEventBridge();
				_ContainerRelationshipAddedEvent += value;
			}
			remove
			{
				_ContainerRelationshipAddedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 14,15,16
		/// </summary>
		private event Application_ContainerRelationshipDeletedEventHandler _ContainerRelationshipDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765445(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		public event Application_ContainerRelationshipDeletedEventHandler ContainerRelationshipDeletedEvent
		{
			add
			{
				CreateEventBridge();
				_ContainerRelationshipDeletedEvent += value;
			}
			remove
			{
				_ContainerRelationshipDeletedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 14,15,16
		/// </summary>
		private event Application_CalloutRelationshipAddedEventHandler _CalloutRelationshipAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769019(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		public event Application_CalloutRelationshipAddedEventHandler CalloutRelationshipAddedEvent
		{
			add
			{
				CreateEventBridge();
				_CalloutRelationshipAddedEvent += value;
			}
			remove
			{
				_CalloutRelationshipAddedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 14,15,16
		/// </summary>
		private event Application_CalloutRelationshipDeletedEventHandler _CalloutRelationshipDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766994(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		public event Application_CalloutRelationshipDeletedEventHandler CalloutRelationshipDeletedEvent
		{
			add
			{
				CreateEventBridge();
				_CalloutRelationshipDeletedEvent += value;
			}
			remove
			{
				_CalloutRelationshipDeletedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 14,15,16
		/// </summary>
		private event Application_RuleSetValidatedEventHandler _RuleSetValidatedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768391(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		public event Application_RuleSetValidatedEventHandler RuleSetValidatedEvent
		{
			add
			{
				CreateEventBridge();
				_RuleSetValidatedEvent += value;
			}
			remove
			{
				_RuleSetValidatedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 15, 16
		/// </summary>
		private event Application_QueryCancelReplaceShapesEventHandler _QueryCancelReplaceShapesEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		public event Application_QueryCancelReplaceShapesEventHandler QueryCancelReplaceShapesEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelReplaceShapesEvent += value;
			}
			remove
			{
				_QueryCancelReplaceShapesEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 15, 16
		/// </summary>
		private event Application_ReplaceShapesCanceledEventHandler _ReplaceShapesCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		public event Application_ReplaceShapesCanceledEventHandler ReplaceShapesCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_ReplaceShapesCanceledEvent += value;
			}
			remove
			{
				_ReplaceShapesCanceledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 15, 16
		/// </summary>
		private event Application_BeforeReplaceShapesEventHandler _BeforeReplaceShapesEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		public event Application_BeforeReplaceShapesEventHandler BeforeReplaceShapesEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeReplaceShapesEvent += value;
			}
			remove
			{
				_BeforeReplaceShapesEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 15, 16
		/// </summary>
		private event Application_AfterReplaceShapesEventHandler _AfterReplaceShapesEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		public event Application_AfterReplaceShapesEventHandler AfterReplaceShapesEvent
		{
			add
			{
				CreateEventBridge();
				_AfterReplaceShapesEvent += value;
			}
			remove
			{
				_AfterReplaceShapesEvent -= value;
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