using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
using NetOffice.Misc;

namespace NetOffice.MSProjectApi
{

	#region Delegates

	#pragma warning disable
	public delegate void Application_NewProjectEventHandler(NetOffice.MSProjectApi.Project pj);
	public delegate void Application_ProjectBeforeTaskDeleteEventHandler(NetOffice.MSProjectApi.Task tsk, ref bool Cancel);
	public delegate void Application_ProjectBeforeResourceDeleteEventHandler(NetOffice.MSProjectApi.Resource res, ref bool Cancel);
	public delegate void Application_ProjectBeforeAssignmentDeleteEventHandler(NetOffice.MSProjectApi.Assignment asg, ref bool Cancel);
	public delegate void Application_ProjectBeforeTaskChangeEventHandler(NetOffice.MSProjectApi.Task tsk, NetOffice.MSProjectApi.Enums.PjField Field, object NewVal, ref bool Cancel);
	public delegate void Application_ProjectBeforeResourceChangeEventHandler(NetOffice.MSProjectApi.Resource res, NetOffice.MSProjectApi.Enums.PjField Field, object NewVal, ref bool Cancel);
	public delegate void Application_ProjectBeforeAssignmentChangeEventHandler(NetOffice.MSProjectApi.Assignment asg, NetOffice.MSProjectApi.Enums.PjAssignmentField Field, object NewVal, ref bool Cancel);
	public delegate void Application_ProjectBeforeTaskNewEventHandler(NetOffice.MSProjectApi.Project pj, ref bool Cancel);
	public delegate void Application_ProjectBeforeResourceNewEventHandler(NetOffice.MSProjectApi.Project pj, ref bool Cancel);
	public delegate void Application_ProjectBeforeAssignmentNewEventHandler(NetOffice.MSProjectApi.Project pj, ref bool Cancel);
	public delegate void Application_ProjectBeforeCloseEventHandler(NetOffice.MSProjectApi.Project pj, ref bool Cancel);
	public delegate void Application_ProjectBeforePrintEventHandler(NetOffice.MSProjectApi.Project pj, ref bool Cancel);
	public delegate void Application_ProjectBeforeSaveEventHandler(NetOffice.MSProjectApi.Project pj, bool SaveAsUi, ref bool Cancel);
	public delegate void Application_ProjectCalculateEventHandler(NetOffice.MSProjectApi.Project pj);
	public delegate void Application_WindowGoalAreaChangeEventHandler(NetOffice.MSProjectApi.Window Window, Int32 goalArea);
	public delegate void Application_WindowSelectionChangeEventHandler(NetOffice.MSProjectApi.Window Window, NetOffice.MSProjectApi.Selection sel, object selType);
	public delegate void Application_WindowBeforeViewChangeEventHandler(NetOffice.MSProjectApi.Window Window, NetOffice.MSProjectApi.View prevView, NetOffice.MSProjectApi.View newView, bool projectHasViewWindow, NetOffice.MSProjectApi.EventInfo Info);
	public delegate void Application_WindowViewChangeEventHandler(NetOffice.MSProjectApi.Window Window, NetOffice.MSProjectApi.View prevView, NetOffice.MSProjectApi.View newView, bool success);
	public delegate void Application_WindowActivateEventHandler(NetOffice.MSProjectApi.Window activatedWindow);
	public delegate void Application_WindowDeactivateEventHandler(NetOffice.MSProjectApi.Window deactivatedWindow);
	public delegate void Application_WindowSidepaneDisplayChangeEventHandler(NetOffice.MSProjectApi.Window Window, bool Close);
	public delegate void Application_WindowSidepaneTaskChangeEventHandler(NetOffice.MSProjectApi.Window Window, Int32 ID, bool IsGoalArea);
	public delegate void Application_WorkpaneDisplayChangeEventHandler(bool DisplayState);
	public delegate void Application_LoadWebPageEventHandler(NetOffice.MSProjectApi.Window Window, ref string TargetPage);
	public delegate void Application_ProjectAfterSaveEventHandler();
	public delegate void Application_ProjectTaskNewEventHandler(NetOffice.MSProjectApi.Project pj, Int32 ID);
	public delegate void Application_ProjectResourceNewEventHandler(NetOffice.MSProjectApi.Project pj, Int32 ID);
	public delegate void Application_ProjectAssignmentNewEventHandler(NetOffice.MSProjectApi.Project pj, Int32 ID);
	public delegate void Application_ProjectBeforeSaveBaselineEventHandler(NetOffice.MSProjectApi.Project pj, bool Interim, NetOffice.MSProjectApi.Enums.PjBaselines bl, NetOffice.MSProjectApi.Enums.PjSaveBaselineFrom InterimCopy, NetOffice.MSProjectApi.Enums.PjSaveBaselineTo InterimInto, bool AllTasks, bool RollupToSummaryTasks, bool RollupFromSubtasks, NetOffice.MSProjectApi.EventInfo Info);
	public delegate void Application_ProjectBeforeClearBaselineEventHandler(NetOffice.MSProjectApi.Project pj, bool Interim, NetOffice.MSProjectApi.Enums.PjBaselines bl, NetOffice.MSProjectApi.Enums.PjSaveBaselineTo InterimFrom, bool AllTasks, NetOffice.MSProjectApi.EventInfo Info);
	public delegate void Application_ProjectBeforeClose2EventHandler(NetOffice.MSProjectApi.Project pj, NetOffice.MSProjectApi.EventInfo Info);
	public delegate void Application_ProjectBeforePrint2EventHandler(NetOffice.MSProjectApi.Project pj, NetOffice.MSProjectApi.EventInfo Info);
	public delegate void Application_ProjectBeforeSave2EventHandler(NetOffice.MSProjectApi.Project pj, bool SaveAsUi, NetOffice.MSProjectApi.EventInfo Info);
	public delegate void Application_ProjectBeforeTaskDelete2EventHandler(NetOffice.MSProjectApi.Task tsk, NetOffice.MSProjectApi.EventInfo Info);
	public delegate void Application_ProjectBeforeResourceDelete2EventHandler(NetOffice.MSProjectApi.Resource res, NetOffice.MSProjectApi.EventInfo Info);
	public delegate void Application_ProjectBeforeAssignmentDelete2EventHandler(NetOffice.MSProjectApi.Assignment asg, NetOffice.MSProjectApi.EventInfo Info);
	public delegate void Application_ProjectBeforeTaskChange2EventHandler(NetOffice.MSProjectApi.Task tsk, NetOffice.MSProjectApi.Enums.PjField Field, object NewVal, NetOffice.MSProjectApi.EventInfo Info);
	public delegate void Application_ProjectBeforeResourceChange2EventHandler(NetOffice.MSProjectApi.Resource res, NetOffice.MSProjectApi.Enums.PjField Field, object NewVal, NetOffice.MSProjectApi.EventInfo Info);
	public delegate void Application_ProjectBeforeAssignmentChange2EventHandler(NetOffice.MSProjectApi.Assignment asg, NetOffice.MSProjectApi.Enums.PjAssignmentField Field, object NewVal, NetOffice.MSProjectApi.EventInfo Info);
	public delegate void Application_ProjectBeforeTaskNew2EventHandler(NetOffice.MSProjectApi.Project pj, NetOffice.MSProjectApi.EventInfo Info);
	public delegate void Application_ProjectBeforeResourceNew2EventHandler(NetOffice.MSProjectApi.Project pj, NetOffice.MSProjectApi.EventInfo Info);
	public delegate void Application_ProjectBeforeAssignmentNew2EventHandler(NetOffice.MSProjectApi.Project pj, NetOffice.MSProjectApi.EventInfo Info);
	public delegate void Application_ApplicationBeforeCloseEventHandler(NetOffice.MSProjectApi.EventInfo Info);
	public delegate void Application_OnUndoOrRedoEventHandler(string bstrLabel, string bstrGUID, bool fUndo);
	public delegate void Application_AfterCubeBuiltEventHandler(ref string CubeFileName);
	public delegate void Application_LoadWebPaneEventHandler(NetOffice.MSProjectApi.Window Window, ref string TargetPage);
	public delegate void Application_JobStartEventHandler(string bstrName, string bstrprojGuid, string bstrjobGuid, Int32 jobType, Int32 lResult);
	public delegate void Application_JobCompletedEventHandler(string bstrName, string bstrprojGuid, string bstrjobGuid, Int32 jobType, Int32 lResult);
	public delegate void Application_SaveStartingToServerEventHandler(string bstrName, string bstrprojGuid);
	public delegate void Application_SaveCompletedToServerEventHandler(string bstrName, string bstrprojGuid);
	public delegate void Application_ProjectBeforePublishEventHandler(NetOffice.MSProjectApi.Project pj, ref bool Cancel);
	public delegate void Application_PaneActivateEventHandler();
	public delegate void Application_SecondaryViewChangeEventHandler(NetOffice.MSProjectApi.Window Window, NetOffice.MSProjectApi.View prevView, NetOffice.MSProjectApi.View newView, bool success);
	public delegate void Application_IsFunctionalitySupportedEventHandler(string bstrFunctionality, NetOffice.MSProjectApi.EventInfo Info);
	public delegate void Application_ConnectionStatusChangedEventHandler(bool online);
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass Application 
	/// SupportByVersion MSProject, 11,12,14
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff920542(v=office.14).aspx
	///</summary>
	[SupportByVersionAttribute("MSProject", 11,12,14)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class Application : _MSProject,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		_EProjectApp2_SinkHelper __EProjectApp2_SinkHelper;
	
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
		public Application():base("MSProject.Application")
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
        /// Returns all running MSProject.Application objects from the environment/system
        /// </summary>
        /// <returns>an MSProject.Application array</returns>
		public static NetOffice.MSProjectApi.Application[] GetActiveInstances()
		{		
			IDisposableEnumeration proxyList = NetOffice.ProxyService.GetActiveInstances("MSProject","Application");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.MSProjectApi.Application> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.MSProjectApi.Application>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.MSProjectApi.Application(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// Returns a running MSProject.Application object from the environment/system.
        /// </summary>
        /// <returns>an MSProject.Application object or null</returns>
		public static NetOffice.MSProjectApi.Application GetActiveInstance()
		{
			object proxy  = NetOffice.ProxyService.GetActiveInstance("MSProject","Application", false);
			if(null != proxy)
				return new NetOffice.MSProjectApi.Application(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// Returns a running MSProject.Application object from the environment/system. 
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an MSProject.Application object or null</returns>
		public static NetOffice.MSProjectApi.Application GetActiveInstance(bool throwOnError)
		{
			object proxy  = NetOffice.ProxyService.GetActiveInstance("MSProject","Application", throwOnError);
			if(null != proxy)
				return new NetOffice.MSProjectApi.Application(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_NewProjectEventHandler _NewProjectEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_NewProjectEventHandler NewProjectEvent
		{
			add
			{
				CreateEventBridge();
				_NewProjectEvent += value;
			}
			remove
			{
				_NewProjectEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectBeforeTaskDeleteEventHandler _ProjectBeforeTaskDeleteEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectBeforeTaskDeleteEventHandler ProjectBeforeTaskDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_ProjectBeforeTaskDeleteEvent += value;
			}
			remove
			{
				_ProjectBeforeTaskDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectBeforeResourceDeleteEventHandler _ProjectBeforeResourceDeleteEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectBeforeResourceDeleteEventHandler ProjectBeforeResourceDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_ProjectBeforeResourceDeleteEvent += value;
			}
			remove
			{
				_ProjectBeforeResourceDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectBeforeAssignmentDeleteEventHandler _ProjectBeforeAssignmentDeleteEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectBeforeAssignmentDeleteEventHandler ProjectBeforeAssignmentDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_ProjectBeforeAssignmentDeleteEvent += value;
			}
			remove
			{
				_ProjectBeforeAssignmentDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectBeforeTaskChangeEventHandler _ProjectBeforeTaskChangeEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectBeforeTaskChangeEventHandler ProjectBeforeTaskChangeEvent
		{
			add
			{
				CreateEventBridge();
				_ProjectBeforeTaskChangeEvent += value;
			}
			remove
			{
				_ProjectBeforeTaskChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectBeforeResourceChangeEventHandler _ProjectBeforeResourceChangeEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectBeforeResourceChangeEventHandler ProjectBeforeResourceChangeEvent
		{
			add
			{
				CreateEventBridge();
				_ProjectBeforeResourceChangeEvent += value;
			}
			remove
			{
				_ProjectBeforeResourceChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectBeforeAssignmentChangeEventHandler _ProjectBeforeAssignmentChangeEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectBeforeAssignmentChangeEventHandler ProjectBeforeAssignmentChangeEvent
		{
			add
			{
				CreateEventBridge();
				_ProjectBeforeAssignmentChangeEvent += value;
			}
			remove
			{
				_ProjectBeforeAssignmentChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectBeforeTaskNewEventHandler _ProjectBeforeTaskNewEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectBeforeTaskNewEventHandler ProjectBeforeTaskNewEvent
		{
			add
			{
				CreateEventBridge();
				_ProjectBeforeTaskNewEvent += value;
			}
			remove
			{
				_ProjectBeforeTaskNewEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectBeforeResourceNewEventHandler _ProjectBeforeResourceNewEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectBeforeResourceNewEventHandler ProjectBeforeResourceNewEvent
		{
			add
			{
				CreateEventBridge();
				_ProjectBeforeResourceNewEvent += value;
			}
			remove
			{
				_ProjectBeforeResourceNewEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectBeforeAssignmentNewEventHandler _ProjectBeforeAssignmentNewEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectBeforeAssignmentNewEventHandler ProjectBeforeAssignmentNewEvent
		{
			add
			{
				CreateEventBridge();
				_ProjectBeforeAssignmentNewEvent += value;
			}
			remove
			{
				_ProjectBeforeAssignmentNewEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectBeforeCloseEventHandler _ProjectBeforeCloseEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectBeforeCloseEventHandler ProjectBeforeCloseEvent
		{
			add
			{
				CreateEventBridge();
				_ProjectBeforeCloseEvent += value;
			}
			remove
			{
				_ProjectBeforeCloseEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectBeforePrintEventHandler _ProjectBeforePrintEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectBeforePrintEventHandler ProjectBeforePrintEvent
		{
			add
			{
				CreateEventBridge();
				_ProjectBeforePrintEvent += value;
			}
			remove
			{
				_ProjectBeforePrintEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectBeforeSaveEventHandler _ProjectBeforeSaveEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectBeforeSaveEventHandler ProjectBeforeSaveEvent
		{
			add
			{
				CreateEventBridge();
				_ProjectBeforeSaveEvent += value;
			}
			remove
			{
				_ProjectBeforeSaveEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectCalculateEventHandler _ProjectCalculateEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectCalculateEventHandler ProjectCalculateEvent
		{
			add
			{
				CreateEventBridge();
				_ProjectCalculateEvent += value;
			}
			remove
			{
				_ProjectCalculateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_WindowGoalAreaChangeEventHandler _WindowGoalAreaChangeEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_WindowGoalAreaChangeEventHandler WindowGoalAreaChangeEvent
		{
			add
			{
				CreateEventBridge();
				_WindowGoalAreaChangeEvent += value;
			}
			remove
			{
				_WindowGoalAreaChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_WindowSelectionChangeEventHandler _WindowSelectionChangeEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_WindowBeforeViewChangeEventHandler _WindowBeforeViewChangeEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_WindowBeforeViewChangeEventHandler WindowBeforeViewChangeEvent
		{
			add
			{
				CreateEventBridge();
				_WindowBeforeViewChangeEvent += value;
			}
			remove
			{
				_WindowBeforeViewChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_WindowViewChangeEventHandler _WindowViewChangeEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_WindowViewChangeEventHandler WindowViewChangeEvent
		{
			add
			{
				CreateEventBridge();
				_WindowViewChangeEvent += value;
			}
			remove
			{
				_WindowViewChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_WindowActivateEventHandler _WindowActivateEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_WindowDeactivateEventHandler _WindowDeactivateEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_WindowSidepaneDisplayChangeEventHandler _WindowSidepaneDisplayChangeEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_WindowSidepaneDisplayChangeEventHandler WindowSidepaneDisplayChangeEvent
		{
			add
			{
				CreateEventBridge();
				_WindowSidepaneDisplayChangeEvent += value;
			}
			remove
			{
				_WindowSidepaneDisplayChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_WindowSidepaneTaskChangeEventHandler _WindowSidepaneTaskChangeEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_WindowSidepaneTaskChangeEventHandler WindowSidepaneTaskChangeEvent
		{
			add
			{
				CreateEventBridge();
				_WindowSidepaneTaskChangeEvent += value;
			}
			remove
			{
				_WindowSidepaneTaskChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_WorkpaneDisplayChangeEventHandler _WorkpaneDisplayChangeEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_WorkpaneDisplayChangeEventHandler WorkpaneDisplayChangeEvent
		{
			add
			{
				CreateEventBridge();
				_WorkpaneDisplayChangeEvent += value;
			}
			remove
			{
				_WorkpaneDisplayChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_LoadWebPageEventHandler _LoadWebPageEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_LoadWebPageEventHandler LoadWebPageEvent
		{
			add
			{
				CreateEventBridge();
				_LoadWebPageEvent += value;
			}
			remove
			{
				_LoadWebPageEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectAfterSaveEventHandler _ProjectAfterSaveEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectAfterSaveEventHandler ProjectAfterSaveEvent
		{
			add
			{
				CreateEventBridge();
				_ProjectAfterSaveEvent += value;
			}
			remove
			{
				_ProjectAfterSaveEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectTaskNewEventHandler _ProjectTaskNewEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectTaskNewEventHandler ProjectTaskNewEvent
		{
			add
			{
				CreateEventBridge();
				_ProjectTaskNewEvent += value;
			}
			remove
			{
				_ProjectTaskNewEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectResourceNewEventHandler _ProjectResourceNewEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectResourceNewEventHandler ProjectResourceNewEvent
		{
			add
			{
				CreateEventBridge();
				_ProjectResourceNewEvent += value;
			}
			remove
			{
				_ProjectResourceNewEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectAssignmentNewEventHandler _ProjectAssignmentNewEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectAssignmentNewEventHandler ProjectAssignmentNewEvent
		{
			add
			{
				CreateEventBridge();
				_ProjectAssignmentNewEvent += value;
			}
			remove
			{
				_ProjectAssignmentNewEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectBeforeSaveBaselineEventHandler _ProjectBeforeSaveBaselineEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectBeforeSaveBaselineEventHandler ProjectBeforeSaveBaselineEvent
		{
			add
			{
				CreateEventBridge();
				_ProjectBeforeSaveBaselineEvent += value;
			}
			remove
			{
				_ProjectBeforeSaveBaselineEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectBeforeClearBaselineEventHandler _ProjectBeforeClearBaselineEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectBeforeClearBaselineEventHandler ProjectBeforeClearBaselineEvent
		{
			add
			{
				CreateEventBridge();
				_ProjectBeforeClearBaselineEvent += value;
			}
			remove
			{
				_ProjectBeforeClearBaselineEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectBeforeClose2EventHandler _ProjectBeforeClose2Event;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectBeforeClose2EventHandler ProjectBeforeClose2Event
		{
			add
			{
				CreateEventBridge();
				_ProjectBeforeClose2Event += value;
			}
			remove
			{
				_ProjectBeforeClose2Event -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectBeforePrint2EventHandler _ProjectBeforePrint2Event;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectBeforePrint2EventHandler ProjectBeforePrint2Event
		{
			add
			{
				CreateEventBridge();
				_ProjectBeforePrint2Event += value;
			}
			remove
			{
				_ProjectBeforePrint2Event -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectBeforeSave2EventHandler _ProjectBeforeSave2Event;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectBeforeSave2EventHandler ProjectBeforeSave2Event
		{
			add
			{
				CreateEventBridge();
				_ProjectBeforeSave2Event += value;
			}
			remove
			{
				_ProjectBeforeSave2Event -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectBeforeTaskDelete2EventHandler _ProjectBeforeTaskDelete2Event;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectBeforeTaskDelete2EventHandler ProjectBeforeTaskDelete2Event
		{
			add
			{
				CreateEventBridge();
				_ProjectBeforeTaskDelete2Event += value;
			}
			remove
			{
				_ProjectBeforeTaskDelete2Event -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectBeforeResourceDelete2EventHandler _ProjectBeforeResourceDelete2Event;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectBeforeResourceDelete2EventHandler ProjectBeforeResourceDelete2Event
		{
			add
			{
				CreateEventBridge();
				_ProjectBeforeResourceDelete2Event += value;
			}
			remove
			{
				_ProjectBeforeResourceDelete2Event -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectBeforeAssignmentDelete2EventHandler _ProjectBeforeAssignmentDelete2Event;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectBeforeAssignmentDelete2EventHandler ProjectBeforeAssignmentDelete2Event
		{
			add
			{
				CreateEventBridge();
				_ProjectBeforeAssignmentDelete2Event += value;
			}
			remove
			{
				_ProjectBeforeAssignmentDelete2Event -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectBeforeTaskChange2EventHandler _ProjectBeforeTaskChange2Event;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectBeforeTaskChange2EventHandler ProjectBeforeTaskChange2Event
		{
			add
			{
				CreateEventBridge();
				_ProjectBeforeTaskChange2Event += value;
			}
			remove
			{
				_ProjectBeforeTaskChange2Event -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectBeforeResourceChange2EventHandler _ProjectBeforeResourceChange2Event;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectBeforeResourceChange2EventHandler ProjectBeforeResourceChange2Event
		{
			add
			{
				CreateEventBridge();
				_ProjectBeforeResourceChange2Event += value;
			}
			remove
			{
				_ProjectBeforeResourceChange2Event -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectBeforeAssignmentChange2EventHandler _ProjectBeforeAssignmentChange2Event;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectBeforeAssignmentChange2EventHandler ProjectBeforeAssignmentChange2Event
		{
			add
			{
				CreateEventBridge();
				_ProjectBeforeAssignmentChange2Event += value;
			}
			remove
			{
				_ProjectBeforeAssignmentChange2Event -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectBeforeTaskNew2EventHandler _ProjectBeforeTaskNew2Event;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectBeforeTaskNew2EventHandler ProjectBeforeTaskNew2Event
		{
			add
			{
				CreateEventBridge();
				_ProjectBeforeTaskNew2Event += value;
			}
			remove
			{
				_ProjectBeforeTaskNew2Event -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectBeforeResourceNew2EventHandler _ProjectBeforeResourceNew2Event;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectBeforeResourceNew2EventHandler ProjectBeforeResourceNew2Event
		{
			add
			{
				CreateEventBridge();
				_ProjectBeforeResourceNew2Event += value;
			}
			remove
			{
				_ProjectBeforeResourceNew2Event -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectBeforeAssignmentNew2EventHandler _ProjectBeforeAssignmentNew2Event;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectBeforeAssignmentNew2EventHandler ProjectBeforeAssignmentNew2Event
		{
			add
			{
				CreateEventBridge();
				_ProjectBeforeAssignmentNew2Event += value;
			}
			remove
			{
				_ProjectBeforeAssignmentNew2Event -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ApplicationBeforeCloseEventHandler _ApplicationBeforeCloseEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ApplicationBeforeCloseEventHandler ApplicationBeforeCloseEvent
		{
			add
			{
				CreateEventBridge();
				_ApplicationBeforeCloseEvent += value;
			}
			remove
			{
				_ApplicationBeforeCloseEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_OnUndoOrRedoEventHandler _OnUndoOrRedoEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_OnUndoOrRedoEventHandler OnUndoOrRedoEvent
		{
			add
			{
				CreateEventBridge();
				_OnUndoOrRedoEvent += value;
			}
			remove
			{
				_OnUndoOrRedoEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_AfterCubeBuiltEventHandler _AfterCubeBuiltEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_AfterCubeBuiltEventHandler AfterCubeBuiltEvent
		{
			add
			{
				CreateEventBridge();
				_AfterCubeBuiltEvent += value;
			}
			remove
			{
				_AfterCubeBuiltEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_LoadWebPaneEventHandler _LoadWebPaneEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_LoadWebPaneEventHandler LoadWebPaneEvent
		{
			add
			{
				CreateEventBridge();
				_LoadWebPaneEvent += value;
			}
			remove
			{
				_LoadWebPaneEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_JobStartEventHandler _JobStartEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_JobStartEventHandler JobStartEvent
		{
			add
			{
				CreateEventBridge();
				_JobStartEvent += value;
			}
			remove
			{
				_JobStartEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_JobCompletedEventHandler _JobCompletedEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_JobCompletedEventHandler JobCompletedEvent
		{
			add
			{
				CreateEventBridge();
				_JobCompletedEvent += value;
			}
			remove
			{
				_JobCompletedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_SaveStartingToServerEventHandler _SaveStartingToServerEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_SaveStartingToServerEventHandler SaveStartingToServerEvent
		{
			add
			{
				CreateEventBridge();
				_SaveStartingToServerEvent += value;
			}
			remove
			{
				_SaveStartingToServerEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_SaveCompletedToServerEventHandler _SaveCompletedToServerEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_SaveCompletedToServerEventHandler SaveCompletedToServerEvent
		{
			add
			{
				CreateEventBridge();
				_SaveCompletedToServerEvent += value;
			}
			remove
			{
				_SaveCompletedToServerEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ProjectBeforePublishEventHandler _ProjectBeforePublishEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ProjectBeforePublishEventHandler ProjectBeforePublishEvent
		{
			add
			{
				CreateEventBridge();
				_ProjectBeforePublishEvent += value;
			}
			remove
			{
				_ProjectBeforePublishEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_PaneActivateEventHandler _PaneActivateEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_PaneActivateEventHandler PaneActivateEvent
		{
			add
			{
				CreateEventBridge();
				_PaneActivateEvent += value;
			}
			remove
			{
				_PaneActivateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_SecondaryViewChangeEventHandler _SecondaryViewChangeEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_SecondaryViewChangeEventHandler SecondaryViewChangeEvent
		{
			add
			{
				CreateEventBridge();
				_SecondaryViewChangeEvent += value;
			}
			remove
			{
				_SecondaryViewChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_IsFunctionalitySupportedEventHandler _IsFunctionalitySupportedEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_IsFunctionalitySupportedEventHandler IsFunctionalitySupportedEvent
		{
			add
			{
				CreateEventBridge();
				_IsFunctionalitySupportedEvent += value;
			}
			remove
			{
				_IsFunctionalitySupportedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Application_ConnectionStatusChangedEventHandler _ConnectionStatusChangedEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public event Application_ConnectionStatusChangedEventHandler ConnectionStatusChangedEvent
		{
			add
			{
				CreateEventBridge();
				_ConnectionStatusChangedEvent += value;
			}
			remove
			{
				_ConnectionStatusChangedEvent -= value;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, _EProjectApp2_SinkHelper.Id);


			if(_EProjectApp2_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				__EProjectApp2_SinkHelper = new _EProjectApp2_SinkHelper(this, _connectPoint);
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
			if( null != __EProjectApp2_SinkHelper)
			{
				__EProjectApp2_SinkHelper.Dispose();
				__EProjectApp2_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}