using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.MSProjectApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Application_NewProjectEventHandler(NetOffice.MSProjectApi.Project pj);
	public delegate void Application_ProjectBeforeTaskDeleteEventHandler(NetOffice.MSProjectApi.Task tsk, ref bool cancel);
	public delegate void Application_ProjectBeforeResourceDeleteEventHandler(NetOffice.MSProjectApi.Resource res, ref bool cancel);
	public delegate void Application_ProjectBeforeAssignmentDeleteEventHandler(NetOffice.MSProjectApi.Assignment asg, ref bool cancel);
	public delegate void Application_ProjectBeforeTaskChangeEventHandler(NetOffice.MSProjectApi.Task tsk, NetOffice.MSProjectApi.Enums.PjField field, object newVal, ref bool cancel);
	public delegate void Application_ProjectBeforeResourceChangeEventHandler(NetOffice.MSProjectApi.Resource res, NetOffice.MSProjectApi.Enums.PjField field, object newVal, ref bool cancel);
	public delegate void Application_ProjectBeforeAssignmentChangeEventHandler(NetOffice.MSProjectApi.Assignment asg, NetOffice.MSProjectApi.Enums.PjAssignmentField field, object newVal, ref bool cancel);
	public delegate void Application_ProjectBeforeTaskNewEventHandler(NetOffice.MSProjectApi.Project pj, ref bool cancel);
	public delegate void Application_ProjectBeforeResourceNewEventHandler(NetOffice.MSProjectApi.Project pj, ref bool cancel);
	public delegate void Application_ProjectBeforeAssignmentNewEventHandler(NetOffice.MSProjectApi.Project pj, ref bool cancel);
	public delegate void Application_ProjectBeforeCloseEventHandler(NetOffice.MSProjectApi.Project pj, ref bool cancel);
	public delegate void Application_ProjectBeforePrintEventHandler(NetOffice.MSProjectApi.Project pj, ref bool cancel);
	public delegate void Application_ProjectBeforeSaveEventHandler(NetOffice.MSProjectApi.Project pj, bool saveAsUi, ref bool cancel);
	public delegate void Application_ProjectCalculateEventHandler(NetOffice.MSProjectApi.Project pj);
	public delegate void Application_WindowGoalAreaChangeEventHandler(NetOffice.MSProjectApi.Window window, Int32 goalArea);
	public delegate void Application_WindowSelectionChangeEventHandler(NetOffice.MSProjectApi.Window window, NetOffice.MSProjectApi.Selection sel, object selType);
	public delegate void Application_WindowBeforeViewChangeEventHandler(NetOffice.MSProjectApi.Window window, NetOffice.MSProjectApi.View prevView, NetOffice.MSProjectApi.View newView, bool projectHasViewWindow, NetOffice.MSProjectApi.EventInfo info);
	public delegate void Application_WindowViewChangeEventHandler(NetOffice.MSProjectApi.Window window, NetOffice.MSProjectApi.View prevView, NetOffice.MSProjectApi.View newView, bool success);
	public delegate void Application_WindowActivateEventHandler(NetOffice.MSProjectApi.Window activatedWindow);
	public delegate void Application_WindowDeactivateEventHandler(NetOffice.MSProjectApi.Window deactivatedWindow);
	public delegate void Application_WindowSidepaneDisplayChangeEventHandler(NetOffice.MSProjectApi.Window Window, bool close);
	public delegate void Application_WindowSidepaneTaskChangeEventHandler(NetOffice.MSProjectApi.Window window, Int32 id, bool isGoalArea);
	public delegate void Application_WorkpaneDisplayChangeEventHandler(bool displayState);
	public delegate void Application_LoadWebPageEventHandler(NetOffice.MSProjectApi.Window window, ref string targetPage);
	public delegate void Application_ProjectAfterSaveEventHandler();
	public delegate void Application_ProjectTaskNewEventHandler(NetOffice.MSProjectApi.Project pj, Int32 id);
	public delegate void Application_ProjectResourceNewEventHandler(NetOffice.MSProjectApi.Project pj, Int32 id);
	public delegate void Application_ProjectAssignmentNewEventHandler(NetOffice.MSProjectApi.Project pj, Int32 ID);
	public delegate void Application_ProjectBeforeSaveBaselineEventHandler(NetOffice.MSProjectApi.Project pj, bool interim, NetOffice.MSProjectApi.Enums.PjBaselines bl, NetOffice.MSProjectApi.Enums.PjSaveBaselineFrom InterimCopy, NetOffice.MSProjectApi.Enums.PjSaveBaselineTo InterimInto, bool AllTasks, bool RollupToSummaryTasks, bool RollupFromSubtasks, NetOffice.MSProjectApi.EventInfo Info);
	public delegate void Application_ProjectBeforeClearBaselineEventHandler(NetOffice.MSProjectApi.Project pj, bool interim, NetOffice.MSProjectApi.Enums.PjBaselines bl, NetOffice.MSProjectApi.Enums.PjSaveBaselineTo InterimFrom, bool AllTasks, NetOffice.MSProjectApi.EventInfo Info);
	public delegate void Application_ProjectBeforeClose2EventHandler(NetOffice.MSProjectApi.Project pj, NetOffice.MSProjectApi.EventInfo info);
	public delegate void Application_ProjectBeforePrint2EventHandler(NetOffice.MSProjectApi.Project pj, NetOffice.MSProjectApi.EventInfo info);
	public delegate void Application_ProjectBeforeSave2EventHandler(NetOffice.MSProjectApi.Project pj, bool SaveAsUi, NetOffice.MSProjectApi.EventInfo info);
	public delegate void Application_ProjectBeforeTaskDelete2EventHandler(NetOffice.MSProjectApi.Task tsk, NetOffice.MSProjectApi.EventInfo info);
	public delegate void Application_ProjectBeforeResourceDelete2EventHandler(NetOffice.MSProjectApi.Resource res, NetOffice.MSProjectApi.EventInfo info);
	public delegate void Application_ProjectBeforeAssignmentDelete2EventHandler(NetOffice.MSProjectApi.Assignment asg, NetOffice.MSProjectApi.EventInfo info);
	public delegate void Application_ProjectBeforeTaskChange2EventHandler(NetOffice.MSProjectApi.Task tsk, NetOffice.MSProjectApi.Enums.PjField Field, object newVal, NetOffice.MSProjectApi.EventInfo info);
	public delegate void Application_ProjectBeforeResourceChange2EventHandler(NetOffice.MSProjectApi.Resource res, NetOffice.MSProjectApi.Enums.PjField field, object newVal, NetOffice.MSProjectApi.EventInfo info);
	public delegate void Application_ProjectBeforeAssignmentChange2EventHandler(NetOffice.MSProjectApi.Assignment asg, NetOffice.MSProjectApi.Enums.PjAssignmentField Field, object newVal, NetOffice.MSProjectApi.EventInfo info);
	public delegate void Application_ProjectBeforeTaskNew2EventHandler(NetOffice.MSProjectApi.Project pj, NetOffice.MSProjectApi.EventInfo info);
	public delegate void Application_ProjectBeforeResourceNew2EventHandler(NetOffice.MSProjectApi.Project pj, NetOffice.MSProjectApi.EventInfo info);
	public delegate void Application_ProjectBeforeAssignmentNew2EventHandler(NetOffice.MSProjectApi.Project pj, NetOffice.MSProjectApi.EventInfo info);
	public delegate void Application_ApplicationBeforeCloseEventHandler(NetOffice.MSProjectApi.EventInfo info);
	public delegate void Application_OnUndoOrRedoEventHandler(string bstrLabel, string bstrGUID, bool fUndo);
	public delegate void Application_AfterCubeBuiltEventHandler(ref string cubeFileName);
	public delegate void Application_LoadWebPaneEventHandler(NetOffice.MSProjectApi.Window window, ref string targetPage);
	public delegate void Application_JobStartEventHandler(string bstrName, string bstrprojGuid, string bstrjobGuid, Int32 jobType, Int32 lResult);
	public delegate void Application_JobCompletedEventHandler(string bstrName, string bstrprojGuid, string bstrjobGuid, Int32 jobType, Int32 lResult);
	public delegate void Application_SaveStartingToServerEventHandler(string bstrName, string bstrprojGuid);
	public delegate void Application_SaveCompletedToServerEventHandler(string bstrName, string bstrprojGuid);
	public delegate void Application_ProjectBeforePublishEventHandler(NetOffice.MSProjectApi.Project pj, ref bool cancel);
	public delegate void Application_PaneActivateEventHandler();
	public delegate void Application_SecondaryViewChangeEventHandler(NetOffice.MSProjectApi.Window Window, NetOffice.MSProjectApi.View prevView, NetOffice.MSProjectApi.View newView, bool success);
	public delegate void Application_IsFunctionalitySupportedEventHandler(string bstrFunctionality, NetOffice.MSProjectApi.EventInfo info);
	public delegate void Application_ConnectionStatusChangedEventHandler(bool online);
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass Application
    /// This class is an alias/typedef for NetOffice.MSProject.Behind.Application
    /// SupportByVersion MSProject 9, 10, 11, 12, 14, 15, 16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194565.aspx </remarks>
    [SupportByVersion("MSProject", 11, 12, 14)]
    [EntityType(EntityType.IsCoClass)]
    [InteropCompatibilityClass]
    public class ApplicationClass : NetOffice.MSProjectApi.Behind.Application
    {
        private string _defaultProgId = "MSProject.Application";

        /// <summary>
        /// Creates a new instance of Microsoft Project
        /// </summary>
        public ApplicationClass()
        {
            ICOMObjectInitialize init = (ICOMObjectInitialize)this;
            init.InitializeCOMObject(_defaultProgId);
        }

        /// <summary>
        /// Creates a new instance of Microsoft Project based on given id.
        /// This can be used to target a specific version of Microsoft Project.
        /// Example usage:
        /// "Microsoft.MSProject.12" to target MSProject 2007
        /// "Microsoft.MSProject.14" to target MSProject 2010
        /// </summary>
        /// <param name="progId">given progid for specific version</param>
        public ApplicationClass(string progId)
        {
            ICOMObjectInitialize init = (ICOMObjectInitialize)this;
            init.InitializeCOMObject(progId);
        }

        /// <summary>
        /// Try get accessing a running application or create a new instance of Microsoft Project
        /// <param name="factory">factory core instead of default core</param>
        /// <param name="tryProxyServiceFirst">try to get a running application first before create a new application</param>
        /// </summary>
        public ApplicationClass(Core factory = null, bool tryProxyServiceFirst = false) : base(factory, tryProxyServiceFirst)
        {

        }

        /// <summary>
        /// Creates a new instance of Microsoft Project
        /// </summary>
        /// <param name="mode">indicates where is the call coming from</param>
        public ApplicationClass(NetOffice.Callers.InteropCompatibilityClassCreateMode mode)
        {
            if (mode == NetOffice.Callers.InteropCompatibilityClassCreateMode.Direct)
            {
                ICOMObjectInitialize init = (ICOMObjectInitialize)this;
                init.InitializeCOMObject(_defaultProgId);
            }
        }
    }

    /// <summary>
    /// CoClass Application
    /// SupportByVersion MSProject, 11,12,14
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920542(v=office.14).aspx </remarks>
    [SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsCoClass), ComProgId("MSProject.Application"), ModuleProvider(typeof(ModulesLegacy.ApplicationModule))]
    [ComEventContract(typeof(EventContracts._EProjectApp2))]
	[TypeId("36D27C48-A1E8-11D3-BA55-00C04F72F325")]
    public interface Application : _MSProject, ICloneable<Application>, IEventBinding, ICOMObjectProxyService
    {
		#region Events

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_NewProjectEventHandler NewProjectEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectBeforeTaskDeleteEventHandler ProjectBeforeTaskDeleteEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectBeforeResourceDeleteEventHandler ProjectBeforeResourceDeleteEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectBeforeAssignmentDeleteEventHandler ProjectBeforeAssignmentDeleteEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectBeforeTaskChangeEventHandler ProjectBeforeTaskChangeEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectBeforeResourceChangeEventHandler ProjectBeforeResourceChangeEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectBeforeAssignmentChangeEventHandler ProjectBeforeAssignmentChangeEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectBeforeTaskNewEventHandler ProjectBeforeTaskNewEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectBeforeResourceNewEventHandler ProjectBeforeResourceNewEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectBeforeAssignmentNewEventHandler ProjectBeforeAssignmentNewEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectBeforeCloseEventHandler ProjectBeforeCloseEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectBeforePrintEventHandler ProjectBeforePrintEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectBeforeSaveEventHandler ProjectBeforeSaveEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectCalculateEventHandler ProjectCalculateEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_WindowGoalAreaChangeEventHandler WindowGoalAreaChangeEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_WindowSelectionChangeEventHandler WindowSelectionChangeEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_WindowBeforeViewChangeEventHandler WindowBeforeViewChangeEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_WindowViewChangeEventHandler WindowViewChangeEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_WindowActivateEventHandler WindowActivateEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_WindowDeactivateEventHandler WindowDeactivateEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_WindowSidepaneDisplayChangeEventHandler WindowSidepaneDisplayChangeEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_WindowSidepaneTaskChangeEventHandler WindowSidepaneTaskChangeEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_WorkpaneDisplayChangeEventHandler WorkpaneDisplayChangeEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_LoadWebPageEventHandler LoadWebPageEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectAfterSaveEventHandler ProjectAfterSaveEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectTaskNewEventHandler ProjectTaskNewEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectResourceNewEventHandler ProjectResourceNewEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectAssignmentNewEventHandler ProjectAssignmentNewEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectBeforeSaveBaselineEventHandler ProjectBeforeSaveBaselineEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectBeforeClearBaselineEventHandler ProjectBeforeClearBaselineEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectBeforeClose2EventHandler ProjectBeforeClose2Event;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectBeforePrint2EventHandler ProjectBeforePrint2Event;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectBeforeSave2EventHandler ProjectBeforeSave2Event;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectBeforeTaskDelete2EventHandler ProjectBeforeTaskDelete2Event;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectBeforeResourceDelete2EventHandler ProjectBeforeResourceDelete2Event;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectBeforeAssignmentDelete2EventHandler ProjectBeforeAssignmentDelete2Event;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectBeforeTaskChange2EventHandler ProjectBeforeTaskChange2Event;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectBeforeResourceChange2EventHandler ProjectBeforeResourceChange2Event;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectBeforeAssignmentChange2EventHandler ProjectBeforeAssignmentChange2Event;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectBeforeTaskNew2EventHandler ProjectBeforeTaskNew2Event;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectBeforeResourceNew2EventHandler ProjectBeforeResourceNew2Event;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectBeforeAssignmentNew2EventHandler ProjectBeforeAssignmentNew2Event;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ApplicationBeforeCloseEventHandler ApplicationBeforeCloseEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_OnUndoOrRedoEventHandler OnUndoOrRedoEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_AfterCubeBuiltEventHandler AfterCubeBuiltEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_LoadWebPaneEventHandler LoadWebPaneEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_JobStartEventHandler JobStartEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_JobCompletedEventHandler JobCompletedEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_SaveStartingToServerEventHandler SaveStartingToServerEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_SaveCompletedToServerEventHandler SaveCompletedToServerEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ProjectBeforePublishEventHandler ProjectBeforePublishEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_PaneActivateEventHandler PaneActivateEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_SecondaryViewChangeEventHandler SecondaryViewChangeEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_IsFunctionalitySupportedEventHandler IsFunctionalitySupportedEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Application_ConnectionStatusChangedEventHandler ConnectionStatusChangedEvent;

		#endregion
	}
}
