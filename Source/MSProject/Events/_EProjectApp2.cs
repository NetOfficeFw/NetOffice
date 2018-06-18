using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.MSProjectApi.EventContracts
{
    /// <summary>
    /// _EProjectApp2
    /// </summary>
	[SupportByVersion("MSProject", 11,12,14)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("5066D7C4-1ED7-48C4-ACE7-299E109D368C"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _EProjectApp2
	{
        /// <summary>
        /// NewProject
        /// </summary>
        /// <param name="pj"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void NewProject([In, MarshalAs(UnmanagedType.IDispatch)] object pj);

        /// <summary>
        /// ProjectBeforeTaskDelete
        /// </summary>
        /// <param name="tsk"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("tsk", typeof(MSProjectApi.Task))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
		void ProjectBeforeTaskDelete([In, MarshalAs(UnmanagedType.IDispatch)] object tsk, [In] [Out] ref object cancel);

        /// <summary>
        /// ProjectBeforeResourceDelete
        /// </summary>
        /// <param name="res"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("res", typeof(MSProjectApi.Resource))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)]
		void ProjectBeforeResourceDelete([In, MarshalAs(UnmanagedType.IDispatch)] object res, [In] [Out] ref object cancel);

        /// <summary>
        /// ProjectBeforeAssignmentDelete
        /// </summary>
        /// <param name="asg"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("asg", typeof(MSProjectApi.Assignment))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)]
		void ProjectBeforeAssignmentDelete([In, MarshalAs(UnmanagedType.IDispatch)] object asg, [In] [Out] ref object cancel);

        /// <summary>
        /// ProjectBeforeTaskChange
        /// </summary>
        /// <param name="tsk"></param>
        /// <param name="field"></param>
        /// <param name="newVal"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("tsk", typeof(MSProjectApi.Task))]
        [SinkArgument("field", SinkArgumentType.Enum, typeof(MSProjectApi.Enums.PjField))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(9)]
		void ProjectBeforeTaskChange([In, MarshalAs(UnmanagedType.IDispatch)] object tsk, [In] object field, [In] object newVal, [In] [Out] ref object cancel);

        /// <summary>
        /// ProjectBeforeResourceChange
        /// </summary>
        /// <param name="res"></param>
        /// <param name="field"></param>
        /// <param name="newVal"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("res", typeof(MSProjectApi.Resource))]
        [SinkArgument("field", SinkArgumentType.Enum, typeof(MSProjectApi.Enums.PjField))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(10)]
		void ProjectBeforeResourceChange([In, MarshalAs(UnmanagedType.IDispatch)] object res, [In] object field, [In] object newVal, [In] [Out] ref object cancel);

        /// <summary>
        /// ProjectBeforeAssignmentChange
        /// </summary>
        /// <param name="asg"></param>
        /// <param name="field"></param>
        /// <param name="newVal"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("asg", typeof(MSProjectApi.Assignment))]
        [SinkArgument("field", SinkArgumentType.Enum, typeof(MSProjectApi.Enums.PjField))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(11)]
		void ProjectBeforeAssignmentChange([In, MarshalAs(UnmanagedType.IDispatch)] object asg, [In] object field, [In] object newVal, [In] [Out] ref object cancel);

        /// <summary>
        /// ProjectBeforeTaskNew
        /// </summary>
        /// <param name="pj"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(12)]
		void ProjectBeforeTaskNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel);

        /// <summary>
        /// ProjectBeforeResourceNew
        /// </summary>
        /// <param name="pj"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(13)]
		void ProjectBeforeResourceNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel);

        /// <summary>
        /// ProjectBeforeAssignmentNew
        /// </summary>
        /// <param name="pj"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(14)]
		void ProjectBeforeAssignmentNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel);

        /// <summary>
        /// ProjectBeforeClose
        /// </summary>
        /// <param name="pj"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void ProjectBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel);

        /// <summary>
        /// ProjectBeforePrint
        /// </summary>
        /// <param name="pj"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void ProjectBeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel);

        /// <summary>
        /// ProjectBeforeSave
        /// </summary>
        /// <param name="pj"></param>
        /// <param name="saveAsUi"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("saveAsUi", SinkArgumentType.Bool)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void ProjectBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object saveAsUi, [In] [Out] ref object cancel);

        /// <summary>
        /// ProjectCalculate
        /// </summary>
        /// <param name="pj"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)]
		void ProjectCalculate([In, MarshalAs(UnmanagedType.IDispatch)] object pj);

        /// <summary>
        /// WindowGoalAreaChange
        /// </summary>
        /// <param name="window"></param>
        /// <param name="goalArea"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("window", typeof(MSProjectApi.Window))]
        [SinkArgument("goalArea", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(15)]
		void WindowGoalAreaChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In] object goalArea);

        /// <summary>
        /// WindowSelectionChange
        /// </summary>
        /// <param name="window"></param>
        /// <param name="sel"></param>
        /// <param name="selType"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("window", typeof(MSProjectApi.Window))]
        [SinkArgument("sel", typeof(MSProjectApi.Selection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16)]
		void WindowSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] object selType);

        /// <summary>
        /// WindowBeforeViewChange
        /// </summary>
        /// <param name="window"></param>
        /// <param name="prevView"></param>
        /// <param name="newView"></param>
        /// <param name="projectHasViewWindow"></param>
        /// <param name="info"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("window", typeof(MSProjectApi.Window))]
        [SinkArgument("prevView", typeof(MSProjectApi.View))]
        [SinkArgument("newView", typeof(MSProjectApi.View))]
        [SinkArgument("projectHasViewWindow", SinkArgumentType.Bool)]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(17)]
		void WindowBeforeViewChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In, MarshalAs(UnmanagedType.IDispatch)] object prevView, [In, MarshalAs(UnmanagedType.IDispatch)] object newView, [In] object projectHasViewWindow, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

        /// <summary>
        /// WindowViewChange
        /// </summary>
        /// <param name="window"></param>
        /// <param name="prevView"></param>
        /// <param name="newView"></param>
        /// <param name="success"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("window", typeof(MSProjectApi.Window))]
        [SinkArgument("prevView", typeof(MSProjectApi.View))]
        [SinkArgument("newView", typeof(MSProjectApi.View))]
        [SinkArgument("success", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(18)]
		void WindowViewChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In, MarshalAs(UnmanagedType.IDispatch)] object prevView, [In, MarshalAs(UnmanagedType.IDispatch)] object newView, [In] object success);

        /// <summary>
        /// WindowActivate
        /// </summary>
        /// <param name="activatedWindow"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("activatedWindow", typeof(MSProjectApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(19)]
		void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object activatedWindow);

        /// <summary>
        /// WindowDeactivate
        /// </summary>
        /// <param name="deactivatedWindow"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("deactivatedWindow", typeof(MSProjectApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(20)]
		void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object deactivatedWindow);

        /// <summary>
        /// WindowSidepaneDisplayChange
        /// </summary>
        /// <param name="window"></param>
        /// <param name="close"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("window", typeof(MSProjectApi.Window))]
        [SinkArgument("close", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(21)]
		void WindowSidepaneDisplayChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In] object close);

        /// <summary>
        /// WindowSidepaneTaskChange
        /// </summary>
        /// <param name="window"></param>
        /// <param name="iD"></param>
        /// <param name="isGoalArea"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("window", typeof(MSProjectApi.Window))]
        [SinkArgument("iD", SinkArgumentType.Int32)]
        [SinkArgument("isGoalArea", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(22)]
		void WindowSidepaneTaskChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In] object iD, [In] object isGoalArea);

        /// <summary>
        /// WorkpaneDisplayChange
        /// </summary>
        /// <param name="displayState"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("displayState", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(23)]
		void WorkpaneDisplayChange([In] object displayState);

        /// <summary>
        /// LoadWebPage
        /// </summary>
        /// <param name="window"></param>
        /// <param name="targetPage"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("window", typeof(MSProjectApi.Window))]
        [SinkArgument("targetPage", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(24)]
		void LoadWebPage([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In] [Out] ref object targetPage);

        /// <summary>
        /// ProjectAfterSave
        /// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(25)]
		void ProjectAfterSave();

        /// <summary>
        /// ProjectTaskNew
        /// </summary>
        /// <param name="pj"></param>
        /// <param name="iD"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("iD", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(26)]
		void ProjectTaskNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object iD);

        /// <summary>
        /// ProjectResourceNew
        /// </summary>
        /// <param name="pj"></param>
        /// <param name="iD"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("iD", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(27)]
		void ProjectResourceNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object iD);

        /// <summary>
        /// ProjectAssignmentNew
        /// </summary>
        /// <param name="pj"></param>
        /// <param name="iD"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("iD", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(28)]
		void ProjectAssignmentNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object iD);

        /// <summary>
        /// ProjectBeforeSaveBaseline
        /// </summary>
        /// <param name="pj"></param>
        /// <param name="interim"></param>
        /// <param name="bl"></param>
        /// <param name="interimCopy"></param>
        /// <param name="interimInto"></param>
        /// <param name="allTasks"></param>
        /// <param name="rollupToSummaryTasks"></param>
        /// <param name="rollupFromSubtasks"></param>
        /// <param name="info"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("interim", SinkArgumentType.Bool)]
        [SinkArgument("bl", SinkArgumentType.Enum, typeof(MSProjectApi.Enums.PjBaselines))]
        [SinkArgument("interimCopy", SinkArgumentType.Enum, typeof(MSProjectApi.Enums.PjSaveBaselineFrom))]
        [SinkArgument("newInterimInto", SinkArgumentType.Enum, typeof(MSProjectApi.Enums.PjSaveBaselineTo))]
        [SinkArgument("allTasks", SinkArgumentType.Bool)]
        [SinkArgument("rollupToSummaryTasks", SinkArgumentType.Bool)]
        [SinkArgument("rollupFromSubtasks", SinkArgumentType.Bool)]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(29)]
		void ProjectBeforeSaveBaseline([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object interim, [In] object bl, [In] object interimCopy, [In] object interimInto, [In] object allTasks, [In] object rollupToSummaryTasks, [In] object rollupFromSubtasks, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

        /// <summary>
        /// ProjectBeforeClearBaseline
        /// </summary>
        /// <param name="pj"></param>
        /// <param name="interim"></param>
        /// <param name="bl"></param>
        /// <param name="interimFrom"></param>
        /// <param name="allTasks"></param>
        /// <param name="info"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("interim", SinkArgumentType.Bool)]
        [SinkArgument("bl", SinkArgumentType.Enum, typeof(MSProjectApi.Enums.PjBaselines))]
        [SinkArgument("interimFrom", SinkArgumentType.Enum, typeof(MSProjectApi.Enums.PjSaveBaselineTo))]
        [SinkArgument("allTasks", SinkArgumentType.Bool)]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(30)]
		void ProjectBeforeClearBaseline([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object interim, [In] object bl, [In] object interimFrom, [In] object allTasks, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

        /// <summary>
        /// ProjectBeforeClose2
        /// </summary>
        /// <param name="pj"></param>
        /// <param name="info"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741826)]
		void ProjectBeforeClose2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

        /// <summary>
        /// ProjectBeforePrint2
        /// </summary>
        /// <param name="pj"></param>
        /// <param name="info"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741828)]
		void ProjectBeforePrint2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

        /// <summary>
        /// ProjectBeforeSave2
        /// </summary>
        /// <param name="pj"></param>
        /// <param name="saveAsUi"></param>
        /// <param name="info"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("saveAsUi", SinkArgumentType.Bool)]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741827)]
		void ProjectBeforeSave2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object saveAsUi, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

        /// <summary>
        /// ProjectBeforeTaskDelete2
        /// </summary>
        /// <param name="tsk"></param>
        /// <param name="info"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("tsk", typeof(MSProjectApi.Task))]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741830)]
		void ProjectBeforeTaskDelete2([In, MarshalAs(UnmanagedType.IDispatch)] object tsk, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

        /// <summary>
        /// ProjectBeforeResourceDelete2
        /// </summary>
        /// <param name="res"></param>
        /// <param name="info"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("res", typeof(MSProjectApi.Resource))]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741831)]
		void ProjectBeforeResourceDelete2([In, MarshalAs(UnmanagedType.IDispatch)] object res, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

        /// <summary>
        /// ProjectBeforeAssignmentDelete2
        /// </summary>
        /// <param name="asg"></param>
        /// <param name="info"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("asg", typeof(MSProjectApi.Assignment))]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741832)]
		void ProjectBeforeAssignmentDelete2([In, MarshalAs(UnmanagedType.IDispatch)] object asg, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

        /// <summary>
        /// ProjectBeforeTaskChange2
        /// </summary>
        /// <param name="tsk"></param>
        /// <param name="field"></param>
        /// <param name="newVal"></param>
        /// <param name="info"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("tsk", typeof(MSProjectApi.Task))]
        [SinkArgument("field", SinkArgumentType.Enum, typeof(MSProjectApi.Enums.PjField))]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741833)]
		void ProjectBeforeTaskChange2([In, MarshalAs(UnmanagedType.IDispatch)] object tsk, [In] object field, [In] object newVal, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

        /// <summary>
        /// ProjectBeforeResourceChange2
        /// </summary>
        /// <param name="res"></param>
        /// <param name="field"></param>
        /// <param name="newVal"></param>
        /// <param name="info"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("res", typeof(MSProjectApi.Resource))]
        [SinkArgument("field", SinkArgumentType.Enum, typeof(MSProjectApi.Enums.PjField))]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741834)]
		void ProjectBeforeResourceChange2([In, MarshalAs(UnmanagedType.IDispatch)] object res, [In] object field, [In] object newVal, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

        /// <summary>
        /// ProjectBeforeAssignmentChange2
        /// </summary>
        /// <param name="asg"></param>
        /// <param name="field"></param>
        /// <param name="newVal"></param>
        /// <param name="info"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("asg", typeof(MSProjectApi.Assignment))]
        [SinkArgument("field", SinkArgumentType.Enum, typeof(MSProjectApi.Enums.PjField))]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741835)]
		void ProjectBeforeAssignmentChange2([In, MarshalAs(UnmanagedType.IDispatch)] object asg, [In] object field, [In] object newVal, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

        /// <summary>
        /// ProjectBeforeTaskNew2
        /// </summary>
        /// <param name="pj"></param>
        /// <param name="info"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741836)]
		void ProjectBeforeTaskNew2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

        /// <summary>
        /// ProjectBeforeResourceNew2
        /// </summary>
        /// <param name="pj"></param>
        /// <param name="info"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741837)]
		void ProjectBeforeResourceNew2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

        /// <summary>
        /// ProjectBeforeAssignmentNew2
        /// </summary>
        /// <param name="pj"></param>
        /// <param name="info"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741838)]
		void ProjectBeforeAssignmentNew2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

        /// <summary>
        /// ApplicationBeforeClose
        /// </summary>
        /// <param name="info"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(31)]
		void ApplicationBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object info);

        /// <summary>
        /// OnUndoOrRedo
        /// </summary>
        /// <param name="bstrLabel"></param>
        /// <param name="bstrGUID"></param>
        /// <param name="fUndo"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("bstrLabel", SinkArgumentType.String)]
        [SinkArgument("bstrGUID", SinkArgumentType.String)]
        [SinkArgument("fUndo", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(32)]
		void OnUndoOrRedo([In] object bstrLabel, [In] object bstrGUID, [In] object fUndo);

        /// <summary>
        /// AfterCubeBuilt
        /// </summary>
        /// <param name="cubeFileName"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("cubeFileName", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(33)]
		void AfterCubeBuilt([In] [Out] ref object cubeFileName);

        /// <summary>
        /// LoadWebPane
        /// </summary>
        /// <param name="window"></param>
        /// <param name="targetPage"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("window", typeof(MSProjectApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(34)]
		void LoadWebPane([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In] [Out] ref object targetPage);

        /// <summary>
        /// JobStart
        /// </summary>
        /// <param name="bstrName"></param>
        /// <param name="bstrprojGuid"></param>
        /// <param name="bstrjobGuid"></param>
        /// <param name="jobType"></param>
        /// <param name="lResult"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("bstrName", SinkArgumentType.String)]
        [SinkArgument("bstrprojGuid", SinkArgumentType.String)]
        [SinkArgument("bstrjobGuid", SinkArgumentType.String)]
        [SinkArgument("jobType", SinkArgumentType.Int32)]
        [SinkArgument("lResult", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(35)]
		void JobStart([In] object bstrName, [In] object bstrprojGuid, [In] object bstrjobGuid, [In] object jobType, [In] object lResult);

        /// <summary>
        /// JobCompleted
        /// </summary>
        /// <param name="bstrName"></param>
        /// <param name="bstrprojGuid"></param>
        /// <param name="bstrjobGuid"></param>
        /// <param name="jobType"></param>
        /// <param name="lResult"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("bstrName", SinkArgumentType.String)]
        [SinkArgument("bstrprojGuid", SinkArgumentType.String)]
        [SinkArgument("bstrjobGuid", SinkArgumentType.String)]
        [SinkArgument("jobType", SinkArgumentType.Int32)]
        [SinkArgument("lResult", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(36)]
		void JobCompleted([In] object bstrName, [In] object bstrprojGuid, [In] object bstrjobGuid, [In] object jobType, [In] object lResult);

        /// <summary>
        /// SaveStartingToServer
        /// </summary>
        /// <param name="bstrName"></param>
        /// <param name="bstrprojGuid"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("bstrName", SinkArgumentType.String)]
        [SinkArgument("bstrprojGuid", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(37)]
		void SaveStartingToServer([In] object bstrName, [In] object bstrprojGuid);

        /// <summary>
        /// SaveCompletedToServer
        /// </summary>
        /// <param name="bstrName"></param>
        /// <param name="bstrprojGuid"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("bstrName", SinkArgumentType.String)]
        [SinkArgument("bstrprojGuid", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(38)]
		void SaveCompletedToServer([In] object bstrName, [In] object bstrprojGuid);

        /// <summary>
        /// ProjectBeforePublish
        /// </summary>
        /// <param name="pj"></param>
        /// <param name="cancel"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(39)]
		void ProjectBeforePublish([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel);

        /// <summary>
        /// PaneActivate
        /// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(40)]
		void PaneActivate();

        /// <summary>
        /// SecondaryViewChange
        /// </summary>
        /// <param name="window"></param>
        /// <param name="prevView"></param>
        /// <param name="newView"></param>
        /// <param name="success"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("window", typeof(MSProjectApi.Window))]
        [SinkArgument("prevView", typeof(MSProjectApi.View))]
        [SinkArgument("newView", typeof(MSProjectApi.View))]
        [SinkArgument("success", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(41)]
		void SecondaryViewChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In, MarshalAs(UnmanagedType.IDispatch)] object prevView, [In, MarshalAs(UnmanagedType.IDispatch)] object newView, [In] object success);

        /// <summary>
        /// IsFunctionalitySupported
        /// </summary>
        /// <param name="bstrFunctionality"></param>
        /// <param name="info"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("bstrFunctionality", SinkArgumentType.String)]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(42)]
		void IsFunctionalitySupported([In] object bstrFunctionality, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

        /// <summary>
        /// ConnectionStatusChanged
        /// </summary>
        /// <param name="online"></param>
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("online", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(43)]
		void ConnectionStatusChanged([In] object online);
	}
}
