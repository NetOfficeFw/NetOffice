using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.MSProjectApi.Behind
{
	/// <summary>
	/// CoClass Application
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920542(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsCoClass), ComProgId("MSProject.Application"), ModuleProvider(typeof(ModulesLegacy.ApplicationModule))]
    [ComEventContract(typeof(NetOffice.MSProjectApi.EventContracts._EProjectApp2))]
    [HasInteropCompatibilityClass(typeof(ApplicationClass))]
    public class Application : _MSProject, NetOffice.MSProjectApi.Application, IAutomaticQuit
    {
		#pragma warning disable

		#region Fields

		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;
        private NetOffice.MSProjectApi.Behind.EventContracts._EProjectApp2_SinkHelper __EProjectApp2_SinkHelper;

		#endregion

		#region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.MSProjectApi.Application);
                return _contractType;
            }
        }
        private static Type _contractType;


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

 	    /// <summary>
        ///Stub Ctor, not intended to use
        /// </summary>
        public Application() : base()
        {

        }

        /// <summary>
        /// Creates a new instance of Application
        /// <param name="enableProxyService">try to get a running application first before create a new application</param>
        /// </summary>
        public Application(Core factory = null, bool enableProxyService = false) : base()
        {
            object proxy = null;
            if (enableProxyService)
            {
                proxy = ProxyService.GetActiveInstance("MSProject", "Application", false);
                if (null != proxy)
                {
                    CreateFromProxy(proxy, true);
                    FromProxyService = true;
                }
            }

            if(null == proxy)
            {
                CreateFromProgId("MSProject.Application", true);
            }

            _callQuitInDispose = null == ParentObject;
            Factory = null != factory ? factory : Core.Default;
            OnCreate();
            ModulesLegacy.ApplicationModule.Instance = this;
        }

        #endregion

        #region ICOMObjectProxyService

        /// <summary>
        /// Instance is created from an already running application
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public virtual bool FromProxyService { get; private set; }

        #endregion

        #region Static CoClass Methods

        /// <summary>
        /// Returns all running MSProject.Application instances from the environment/system
        /// </summary>
        /// <returns>MSProject.Application sequence</returns>
        public static IDisposableSequence<Application> GetActiveInstances()
        {
            return ProxyService.GetActiveInstances<Application>("MSProject", "Application");
        }

        /// <summary>
        /// Returns all running MSProject.Application instances from the environment/system that passed a predicate filter
        /// </summary>
        /// <param name="predicate">filter predicate</param>
        /// <returns>MSProject.Application sequence</returns>
        public static IDisposableSequence<Application> GetActiveInstances(Func<Application, bool> predicate)
        {
            return ProxyService.GetActiveInstances<Application>("MSProject", "Application", predicate);
        }

        /// <summary>
        /// Returns the count of running MSProject.Application instances that passed a predicate filter
        /// </summary>
        /// <returns>count of running application</returns>
        public static int GetActiveInstancesCount()
        {
            var sequence = ProxyService.GetActiveInstances<Application>("MSProject", "Application");
            int result = sequence.Count;
            sequence.Dispose();
            return result;
        }

        /// <summary>
        /// Returns the count of running MSProject.Application instances that passed a predicate filter
        /// </summary>
        /// <param name="predicate">filter predicate</param>
        /// <returns>count of running application</returns>
        public static int GetActiveInstancesCount(Func<Application, bool> predicate)
        {
            var sequence = ProxyService.GetActiveInstances<Application>("MSProject", "Application", predicate);
            int result = sequence.Count;
            sequence.Dispose();
            return result;
        }

        /// <summary>
        /// Returns a running MSProject.Application instance from the environment/system
        /// </summary>
        /// <param name="throwExceptionIfNotFound">throw exception if unable to find an instance</param>
        /// <returns>MSProject.Application instance or null(Nothing in Visual Basic)</returns>
        public static Application GetActiveInstance(bool throwExceptionIfNotFound = false)
        {
            return ProxyService.GetActiveInstance<Application>("MSProject", "Application", throwExceptionIfNotFound);
        }

        /// <summary>
        /// Returns first running MSProject.Application instance from the environment/system that passed a predicate filter
        /// </summary>
        /// <param name="predicate">filter predicate</param>
        /// <param name="throwExceptionIfNotFound">throw exception if unable to find an instance</param>
        /// <returns>MSProject.Application instance or null(Nothing in Visual Basic)</returns>
        /// <exception cref="ArgumentOutOfRangeException">occurs if no instance match and throwExceptionIfNotFound is set</exception>
        public static Application GetActiveInstance(Func<Application, bool> predicate, bool throwExceptionIfNotFound = false)
        {
            return ProxyService.GetActiveInstance<Application>("MSProject", "Application", predicate, throwExceptionIfNotFound);
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
		public virtual event Application_NewProjectEventHandler NewProjectEvent
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
		public virtual event Application_ProjectBeforeTaskDeleteEventHandler ProjectBeforeTaskDeleteEvent
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
		public virtual event Application_ProjectBeforeResourceDeleteEventHandler ProjectBeforeResourceDeleteEvent
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
		public virtual event Application_ProjectBeforeAssignmentDeleteEventHandler ProjectBeforeAssignmentDeleteEvent
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
		public virtual event Application_ProjectBeforeTaskChangeEventHandler ProjectBeforeTaskChangeEvent
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
		public virtual event Application_ProjectBeforeResourceChangeEventHandler ProjectBeforeResourceChangeEvent
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
		public virtual event Application_ProjectBeforeAssignmentChangeEventHandler ProjectBeforeAssignmentChangeEvent
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
		public virtual event Application_ProjectBeforeTaskNewEventHandler ProjectBeforeTaskNewEvent
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
		public virtual event Application_ProjectBeforeResourceNewEventHandler ProjectBeforeResourceNewEvent
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
		public virtual event Application_ProjectBeforeAssignmentNewEventHandler ProjectBeforeAssignmentNewEvent
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
		public virtual event Application_ProjectBeforeCloseEventHandler ProjectBeforeCloseEvent
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
		public virtual event Application_ProjectBeforePrintEventHandler ProjectBeforePrintEvent
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
		public virtual event Application_ProjectBeforeSaveEventHandler ProjectBeforeSaveEvent
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
		public virtual event Application_ProjectCalculateEventHandler ProjectCalculateEvent
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
		public virtual event Application_WindowGoalAreaChangeEventHandler WindowGoalAreaChangeEvent
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
		public virtual event Application_WindowSelectionChangeEventHandler WindowSelectionChangeEvent
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
		public virtual event Application_WindowBeforeViewChangeEventHandler WindowBeforeViewChangeEvent
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
		public virtual event Application_WindowViewChangeEventHandler WindowViewChangeEvent
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
		public virtual event Application_WindowActivateEventHandler WindowActivateEvent
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
		public virtual event Application_WindowDeactivateEventHandler WindowDeactivateEvent
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
		public virtual event Application_WindowSidepaneDisplayChangeEventHandler WindowSidepaneDisplayChangeEvent
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
		public virtual event Application_WindowSidepaneTaskChangeEventHandler WindowSidepaneTaskChangeEvent
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
		public virtual event Application_WorkpaneDisplayChangeEventHandler WorkpaneDisplayChangeEvent
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
		public virtual event Application_LoadWebPageEventHandler LoadWebPageEvent
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
		public virtual event Application_ProjectAfterSaveEventHandler ProjectAfterSaveEvent
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
		public virtual event Application_ProjectTaskNewEventHandler ProjectTaskNewEvent
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
		public virtual event Application_ProjectResourceNewEventHandler ProjectResourceNewEvent
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
		public virtual event Application_ProjectAssignmentNewEventHandler ProjectAssignmentNewEvent
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
		public virtual event Application_ProjectBeforeSaveBaselineEventHandler ProjectBeforeSaveBaselineEvent
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
		public virtual event Application_ProjectBeforeClearBaselineEventHandler ProjectBeforeClearBaselineEvent
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
		public virtual event Application_ProjectBeforeClose2EventHandler ProjectBeforeClose2Event
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
		public virtual event Application_ProjectBeforePrint2EventHandler ProjectBeforePrint2Event
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
		public virtual event Application_ProjectBeforeSave2EventHandler ProjectBeforeSave2Event
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
		public virtual event Application_ProjectBeforeTaskDelete2EventHandler ProjectBeforeTaskDelete2Event
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
		public virtual event Application_ProjectBeforeResourceDelete2EventHandler ProjectBeforeResourceDelete2Event
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
		public virtual event Application_ProjectBeforeAssignmentDelete2EventHandler ProjectBeforeAssignmentDelete2Event
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
		public virtual event Application_ProjectBeforeTaskChange2EventHandler ProjectBeforeTaskChange2Event
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
		public virtual event Application_ProjectBeforeResourceChange2EventHandler ProjectBeforeResourceChange2Event
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
		public virtual event Application_ProjectBeforeAssignmentChange2EventHandler ProjectBeforeAssignmentChange2Event
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
		public virtual event Application_ProjectBeforeTaskNew2EventHandler ProjectBeforeTaskNew2Event
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
		public virtual event Application_ProjectBeforeResourceNew2EventHandler ProjectBeforeResourceNew2Event
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
		public virtual event Application_ProjectBeforeAssignmentNew2EventHandler ProjectBeforeAssignmentNew2Event
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
		public virtual event Application_ApplicationBeforeCloseEventHandler ApplicationBeforeCloseEvent
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
		public virtual event Application_OnUndoOrRedoEventHandler OnUndoOrRedoEvent
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
		public virtual event Application_AfterCubeBuiltEventHandler AfterCubeBuiltEvent
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
		public virtual event Application_LoadWebPaneEventHandler LoadWebPaneEvent
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
		public virtual event Application_JobStartEventHandler JobStartEvent
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
		public virtual event Application_JobCompletedEventHandler JobCompletedEvent
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
		public virtual event Application_SaveStartingToServerEventHandler SaveStartingToServerEvent
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
		public virtual event Application_SaveCompletedToServerEventHandler SaveCompletedToServerEvent
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
		public virtual event Application_ProjectBeforePublishEventHandler ProjectBeforePublishEvent
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
		public virtual event Application_PaneActivateEventHandler PaneActivateEvent
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
		public virtual event Application_SecondaryViewChangeEventHandler SecondaryViewChangeEvent
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
		public virtual event Application_IsFunctionalitySupportedEventHandler IsFunctionalitySupportedEvent
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
		public virtual event Application_ConnectionStatusChangedEventHandler ConnectionStatusChangedEvent
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

        #region IAutomaticQuit

        /// <summary>
        /// Determines Quit method want be called while disposing if NetOffice.Settings.EnableAutomaticQuit is true.
        /// Default is true when instance has no parent object and its not a cloned instance, otherwise false.
        /// </summary>
        bool IAutomaticQuit.Enabled
        {

            get
            {
                return _callQuitInDispose;
            }
            set
            {
                _callQuitInDispose = value;
            }
        }

        #endregion

        #region IEventBinding

        /// <summary>
        /// Creates active sink helper
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual void CreateEventBridge()
        {
			if(false == Factory.Settings.EnableEvents)
				return;

			if (null != _connectPoint)
				return;

            if (null == _activeSinkId)
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, NetOffice.MSProjectApi.Behind.EventContracts._EProjectApp2_SinkHelper.Id);
            

			if(NetOffice.MSProjectApi.Behind.EventContracts._EProjectApp2_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				__EProjectApp2_SinkHelper = new NetOffice.MSProjectApi.Behind.EventContracts._EProjectApp2_SinkHelper(this, _connectPoint);
				return;
			}
        }

        /// <summary>
        /// The instance use currently an event listener
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool EventBridgeInitialized
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
        public virtual bool HasEventRecipients()
        {
            return NetOffice.Events.CoClassEventReflector.HasEventRecipients(this, LateBindingApiWrapperType);
        }

        /// <summary>
        /// Instance has one or more event recipients
        /// </summary>
        /// <param name="eventName">name of the event</param>
        /// <returns></returns>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool HasEventRecipients(string eventName)
        {
            return NetOffice.Events.CoClassEventReflector.HasEventRecipients(this, LateBindingApiWrapperType, eventName);
        }

        /// <summary>
        /// Target methods from its actual event recipients
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Delegate[] GetEventRecipients(string eventName)
        {
            return NetOffice.Events.CoClassEventReflector.GetEventRecipients(this, LateBindingApiWrapperType, eventName);
        }

        /// <summary>
        /// Returns the current count of event recipients
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual int GetCountOfEventRecipients(string eventName)
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
        public virtual int RaiseCustomEvent(string eventName, ref object[] paramsArray)
		{
            return NetOffice.Events.CoClassEventReflector.RaiseCustomEvent(this, LateBindingApiWrapperType, eventName, ref paramsArray);
		}
        /// <summary>
        /// Stop listening events for the instance
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void DisposeEventBridge()
        {
			if( null != __EProjectApp2_SinkHelper)
			{
				__EProjectApp2_SinkHelper.Dispose();
				__EProjectApp2_SinkHelper = null;
			}

			_connectPoint = null;
		}

        #endregion

        #region IDisposable

        /// <summary>
        /// NetOffice method: dispose instance and all child instances
        /// </summary>
        /// <param name="disposeEventBinding">dispose event exported proxies with one or more event recipients</param>
        [Category("NetOffice"), CoreOverridden]
        public override void Dispose(bool disposeEventBinding)
        {
            if (this.Equals(ModulesLegacy.ApplicationModule.Instance))
                ModulesLegacy.ApplicationModule.Instance = null;
            base.Dispose(disposeEventBinding);
        }

        /// <summary>
        /// NetOffice method: dispose instance and all child instances
        /// </summary>
        [Category("NetOffice"), CoreOverridden]
        public override void Dispose()
        {
            if (this.Equals(ModulesLegacy.ApplicationModule.Instance))
                ModulesLegacy.ApplicationModule.Instance = null;
            base.Dispose();
        }

        #endregion

        #region ICloneable<Application>

        /// <summary>
        /// Creates a new Application that is a copy of the current instance
        /// </summary>
        /// <returns>A new Application that is a copy of this instance</returns>
        /// <exception cref="CloneException">An unexpected error occured. See inner exception(s) for details.</exception>
        public new virtual NetOffice.MSProjectApi.Application DeepCopy()
        {
            return base.Clone() as NetOffice.MSProjectApi.Application;
        }

        #endregion

        #pragma warning restore
    }
}
