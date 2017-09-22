using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.MSProjectApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("MSProject", 11,12,14)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("5066D7C4-1ED7-48C4-ACE7-299E109D368C"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _EProjectApp2
	{
		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void NewProject([In, MarshalAs(UnmanagedType.IDispatch)] object pj);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("tsk", typeof(MSProjectApi.Task))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
		void ProjectBeforeTaskDelete([In, MarshalAs(UnmanagedType.IDispatch)] object tsk, [In] [Out] ref object cancel);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("res", typeof(MSProjectApi.Resource))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)]
		void ProjectBeforeResourceDelete([In, MarshalAs(UnmanagedType.IDispatch)] object res, [In] [Out] ref object cancel);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("asg", typeof(MSProjectApi.Assignment))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)]
		void ProjectBeforeAssignmentDelete([In, MarshalAs(UnmanagedType.IDispatch)] object asg, [In] [Out] ref object cancel);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("tsk", typeof(MSProjectApi.Task))]
        [SinkArgument("field", SinkArgumentType.Enum, typeof(MSProjectApi.Enums.PjField))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(9)]
		void ProjectBeforeTaskChange([In, MarshalAs(UnmanagedType.IDispatch)] object tsk, [In] object field, [In] object newVal, [In] [Out] ref object cancel);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("res", typeof(MSProjectApi.Resource))]
        [SinkArgument("field", SinkArgumentType.Enum, typeof(MSProjectApi.Enums.PjField))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(10)]
		void ProjectBeforeResourceChange([In, MarshalAs(UnmanagedType.IDispatch)] object res, [In] object field, [In] object newVal, [In] [Out] ref object cancel);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("asg", typeof(MSProjectApi.Assignment))]
        [SinkArgument("field", SinkArgumentType.Enum, typeof(MSProjectApi.Enums.PjField))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(11)]
		void ProjectBeforeAssignmentChange([In, MarshalAs(UnmanagedType.IDispatch)] object asg, [In] object field, [In] object newVal, [In] [Out] ref object cancel);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(12)]
		void ProjectBeforeTaskNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(13)]
		void ProjectBeforeResourceNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(14)]
		void ProjectBeforeAssignmentNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void ProjectBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void ProjectBeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("saveAsUi", SinkArgumentType.Bool)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void ProjectBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object saveAsUi, [In] [Out] ref object cancel);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)]
		void ProjectCalculate([In, MarshalAs(UnmanagedType.IDispatch)] object pj);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("window", typeof(MSProjectApi.Window))]
        [SinkArgument("goalArea", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(15)]
		void WindowGoalAreaChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In] object goalArea);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("window", typeof(MSProjectApi.Window))]
        [SinkArgument("sel", typeof(MSProjectApi.Selection))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16)]
		void WindowSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] object selType);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("window", typeof(MSProjectApi.Window))]
        [SinkArgument("prevView", typeof(MSProjectApi.View))]
        [SinkArgument("newView", typeof(MSProjectApi.View))]
        [SinkArgument("projectHasViewWindow", SinkArgumentType.Bool)]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(17)]
		void WindowBeforeViewChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In, MarshalAs(UnmanagedType.IDispatch)] object prevView, [In, MarshalAs(UnmanagedType.IDispatch)] object newView, [In] object projectHasViewWindow, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("window", typeof(MSProjectApi.Window))]
        [SinkArgument("prevView", typeof(MSProjectApi.View))]
        [SinkArgument("newView", typeof(MSProjectApi.View))]
        [SinkArgument("success", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(18)]
		void WindowViewChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In, MarshalAs(UnmanagedType.IDispatch)] object prevView, [In, MarshalAs(UnmanagedType.IDispatch)] object newView, [In] object success);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("activatedWindow", typeof(MSProjectApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(19)]
		void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object activatedWindow);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("deactivatedWindow", typeof(MSProjectApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(20)]
		void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object deactivatedWindow);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("window", typeof(MSProjectApi.Window))]
        [SinkArgument("close", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(21)]
		void WindowSidepaneDisplayChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In] object close);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("window", typeof(MSProjectApi.Window))]
        [SinkArgument("iD", SinkArgumentType.Int32)]
        [SinkArgument("isGoalArea", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(22)]
		void WindowSidepaneTaskChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In] object iD, [In] object isGoalArea);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("displayState", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(23)]
		void WorkpaneDisplayChange([In] object displayState);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("window", typeof(MSProjectApi.Window))]
        [SinkArgument("targetPage", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(24)]
		void LoadWebPage([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In] [Out] ref object targetPage);

		[SupportByVersion("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(25)]
		void ProjectAfterSave();

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("iD", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(26)]
		void ProjectTaskNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object iD);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("iD", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(27)]
		void ProjectResourceNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object iD);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("iD", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(28)]
		void ProjectAssignmentNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object iD);

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

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("interim", SinkArgumentType.Bool)]
        [SinkArgument("bl", SinkArgumentType.Enum, typeof(MSProjectApi.Enums.PjBaselines))]
        [SinkArgument("interimFrom", SinkArgumentType.Enum, typeof(MSProjectApi.Enums.PjSaveBaselineTo))]
        [SinkArgument("allTasks", SinkArgumentType.Bool)]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(30)]
		void ProjectBeforeClearBaseline([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object interim, [In] object bl, [In] object interimFrom, [In] object allTasks, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741826)]
		void ProjectBeforeClose2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741828)]
		void ProjectBeforePrint2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("saveAsUi", SinkArgumentType.Bool)]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741827)]
		void ProjectBeforeSave2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object saveAsUi, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("tsk", typeof(MSProjectApi.Task))]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741830)]
		void ProjectBeforeTaskDelete2([In, MarshalAs(UnmanagedType.IDispatch)] object tsk, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("res", typeof(MSProjectApi.Resource))]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741831)]
		void ProjectBeforeResourceDelete2([In, MarshalAs(UnmanagedType.IDispatch)] object res, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("asg", typeof(MSProjectApi.Assignment))]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741832)]
		void ProjectBeforeAssignmentDelete2([In, MarshalAs(UnmanagedType.IDispatch)] object asg, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("tsk", typeof(MSProjectApi.Task))]
        [SinkArgument("field", SinkArgumentType.Enum, typeof(MSProjectApi.Enums.PjField))]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741833)]
		void ProjectBeforeTaskChange2([In, MarshalAs(UnmanagedType.IDispatch)] object tsk, [In] object field, [In] object newVal, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("res", typeof(MSProjectApi.Resource))]
        [SinkArgument("field", SinkArgumentType.Enum, typeof(MSProjectApi.Enums.PjField))]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741834)]
		void ProjectBeforeResourceChange2([In, MarshalAs(UnmanagedType.IDispatch)] object res, [In] object field, [In] object newVal, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("asg", typeof(MSProjectApi.Assignment))]
        [SinkArgument("field", SinkArgumentType.Enum, typeof(MSProjectApi.Enums.PjField))]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741835)]
		void ProjectBeforeAssignmentChange2([In, MarshalAs(UnmanagedType.IDispatch)] object asg, [In] object field, [In] object newVal, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersion("MSProject", 11,12,14)]

        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741836)]
		void ProjectBeforeTaskNew2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741837)]
		void ProjectBeforeResourceNew2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741838)]
		void ProjectBeforeAssignmentNew2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(31)]
		void ApplicationBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("bstrLabel", SinkArgumentType.String)]
        [SinkArgument("bstrGUID", SinkArgumentType.String)]
        [SinkArgument("fUndo", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(32)]
		void OnUndoOrRedo([In] object bstrLabel, [In] object bstrGUID, [In] object fUndo);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("cubeFileName", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(33)]
		void AfterCubeBuilt([In] [Out] ref object cubeFileName);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("window", typeof(MSProjectApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(34)]
		void LoadWebPane([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In] [Out] ref object targetPage);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("bstrName", SinkArgumentType.String)]
        [SinkArgument("bstrprojGuid", SinkArgumentType.String)]
        [SinkArgument("bstrjobGuid", SinkArgumentType.String)]
        [SinkArgument("jobType", SinkArgumentType.Int32)]
        [SinkArgument("lResult", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(35)]
		void JobStart([In] object bstrName, [In] object bstrprojGuid, [In] object bstrjobGuid, [In] object jobType, [In] object lResult);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("bstrName", SinkArgumentType.String)]
        [SinkArgument("bstrprojGuid", SinkArgumentType.String)]
        [SinkArgument("bstrjobGuid", SinkArgumentType.String)]
        [SinkArgument("jobType", SinkArgumentType.Int32)]
        [SinkArgument("lResult", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(36)]
		void JobCompleted([In] object bstrName, [In] object bstrprojGuid, [In] object bstrjobGuid, [In] object jobType, [In] object lResult);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("bstrName", SinkArgumentType.String)]
        [SinkArgument("bstrprojGuid", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(37)]
		void SaveStartingToServer([In] object bstrName, [In] object bstrprojGuid);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("bstrName", SinkArgumentType.String)]
        [SinkArgument("bstrprojGuid", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(38)]
		void SaveCompletedToServer([In] object bstrName, [In] object bstrprojGuid);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("pj", typeof(MSProjectApi.Project))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(39)]
		void ProjectBeforePublish([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel);

		[SupportByVersion("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(40)]
		void PaneActivate();

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("window", typeof(MSProjectApi.Window))]
        [SinkArgument("prevView", typeof(MSProjectApi.View))]
        [SinkArgument("newView", typeof(MSProjectApi.View))]
        [SinkArgument("success", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(41)]
		void SecondaryViewChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In, MarshalAs(UnmanagedType.IDispatch)] object prevView, [In, MarshalAs(UnmanagedType.IDispatch)] object newView, [In] object success);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("bstrFunctionality", SinkArgumentType.String)]
        [SinkArgument("info", typeof(MSProjectApi.EventInfo))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(42)]
		void IsFunctionalitySupported([In] object bstrFunctionality, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersion("MSProject", 11,12,14)]
        [SinkArgument("online", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(43)]
		void ConnectionStatusChanged([In] object online);
	}
	
	#endregion
	
	#region SinkHelper
	
    [InternalEntity(InternalEntityKind.SinkHelper)]
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _EProjectApp2_SinkHelper : SinkHelper, _EProjectApp2
	{
		#region Static
		
		public static readonly string Id = "5066D7C4-1ED7-48C4-ACE7-299E109D368C";
		
		#endregion
	
		#region Ctor

		public _EProjectApp2_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region _EProjectApp2 Members
		
        public void NewProject([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
		{
            if (!Validate("NewProject"))
            {
                Invoker.ReleaseParamsArray(pj);
                return;
            }

			NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			EventBinding.RaiseCustomEvent("NewProject", ref paramsArray);
		}

        public void ProjectBeforeTaskDelete([In, MarshalAs(UnmanagedType.IDispatch)] object tsk, [In] [Out] ref object cancel)
		{
            if (!Validate("ProjectBeforeTaskDelete"))
            {
                Invoker.ReleaseParamsArray(tsk, cancel);
                return;
            }

			NetOffice.MSProjectApi.Task newtsk = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Task>(EventClass, tsk, NetOffice.MSProjectApi.Task.LateBindingApiWrapperType);
			object[] paramsArray = new object[2];
			paramsArray[0] = newtsk;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("ProjectBeforeTaskDelete", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
		}

        public void ProjectBeforeResourceDelete([In, MarshalAs(UnmanagedType.IDispatch)] object res, [In] [Out] ref object cancel)
		{
            if (!Validate("ProjectBeforeResourceDelete"))
            {
                Invoker.ReleaseParamsArray(res, cancel);
                return;
            }

            NetOffice.MSProjectApi.Resource newres = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Resource>(EventClass, res, NetOffice.MSProjectApi.Resource.LateBindingApiWrapperType);
			object[] paramsArray = new object[2];
			paramsArray[0] = newres;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("ProjectBeforeResourceDelete", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
        }

        public void ProjectBeforeAssignmentDelete([In, MarshalAs(UnmanagedType.IDispatch)] object asg, [In] [Out] ref object cancel)
        {
            if (!Validate("ProjectBeforeResourceDelete"))
            {
                Invoker.ReleaseParamsArray(asg, cancel);
                return;
            }

			NetOffice.MSProjectApi.Assignment newasg = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Assignment>(EventClass, asg, NetOffice.MSProjectApi.Assignment.LateBindingApiWrapperType);
			object[] paramsArray = new object[2];
			paramsArray[0] = newasg;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("ProjectBeforeAssignmentDelete", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
        }

        public void ProjectBeforeTaskChange([In, MarshalAs(UnmanagedType.IDispatch)] object tsk, [In] object field, [In] object newVal, [In] [Out] ref object cancel)
		{
            if (!Validate("ProjectBeforeResourceDelete"))
            {
                Invoker.ReleaseParamsArray(tsk, field, newVal, cancel);
                return;
            }

            NetOffice.MSProjectApi.Task newtsk = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Task>(EventClass, tsk, NetOffice.MSProjectApi.Task.LateBindingApiWrapperType);
			NetOffice.MSProjectApi.Enums.PjField newField = (NetOffice.MSProjectApi.Enums.PjField)field;
			object newNewVal = (object)newVal;
			object[] paramsArray = new object[4];
			paramsArray[0] = newtsk;
			paramsArray[1] = newField;
			paramsArray[2] = newNewVal;
			paramsArray.SetValue(cancel, 3);
			EventBinding.RaiseCustomEvent("ProjectBeforeTaskChange", ref paramsArray);

			cancel = ToBoolean(paramsArray[3]);
        }

        public void ProjectBeforeResourceChange([In, MarshalAs(UnmanagedType.IDispatch)] object res, [In] object field, [In] object newVal, [In] [Out] ref object cancel)
        {
            if (!Validate("ProjectBeforeResourceChange"))
            {
                Invoker.ReleaseParamsArray(res, field, newVal, cancel);
                return;
            }

			NetOffice.MSProjectApi.Resource newres = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Resource>(EventClass, res, NetOffice.MSProjectApi.Resource.LateBindingApiWrapperType);
			NetOffice.MSProjectApi.Enums.PjField newField = (NetOffice.MSProjectApi.Enums.PjField)field;
			object newNewVal = (object)newVal;
			object[] paramsArray = new object[4];
			paramsArray[0] = newres;
			paramsArray[1] = newField;
			paramsArray[2] = newNewVal;
			paramsArray.SetValue(cancel, 3);
			EventBinding.RaiseCustomEvent("ProjectBeforeResourceChange", ref paramsArray);

			cancel = ToBoolean(paramsArray[3]);
        }

        public void ProjectBeforeAssignmentChange([In, MarshalAs(UnmanagedType.IDispatch)] object asg, [In] object field, [In] object newVal, [In] [Out] ref object cancel)
		{
            if (!Validate("ProjectBeforeAssignmentChange"))
            {
                Invoker.ReleaseParamsArray(asg, field, newVal, cancel);
                return;
            }

			NetOffice.MSProjectApi.Assignment newasg = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Assignment>(EventClass, asg, NetOffice.MSProjectApi.Assignment.LateBindingApiWrapperType);
			NetOffice.MSProjectApi.Enums.PjAssignmentField newField = (NetOffice.MSProjectApi.Enums.PjAssignmentField)field;
			object newNewVal = (object)newVal;
			object[] paramsArray = new object[4];
			paramsArray[0] = newasg;
			paramsArray[1] = newField;
			paramsArray[2] = newNewVal;
			paramsArray.SetValue(cancel, 3);
			EventBinding.RaiseCustomEvent("ProjectBeforeAssignmentChange", ref paramsArray);

			cancel = ToBoolean(paramsArray[3]);
        }

        public void ProjectBeforeTaskNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel)
		{
            if (!Validate("ProjectBeforeTaskNew"))
            {
                Invoker.ReleaseParamsArray(pj, cancel);
                return;
            }

			NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
			object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("ProjectBeforeTaskNew", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        public void ProjectBeforeResourceNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel)
		{
            if (!Validate("ProjectBeforeResourceNew"))
            {
                Invoker.ReleaseParamsArray(pj, cancel);
                return;
            }

			NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
			object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("ProjectBeforeResourceNew", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
        }

        public void ProjectBeforeAssignmentNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel)
        {
            if (!Validate("ProjectBeforeAssignmentNew"))
            {
                Invoker.ReleaseParamsArray(pj, cancel);
                return;
            }

			NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
			object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("ProjectBeforeAssignmentNew", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
        }

        public void ProjectBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel)
        {
            if (!Validate("ProjectBeforeClose"))
            {
                Invoker.ReleaseParamsArray(pj, cancel);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("ProjectBeforeClose", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        public void ProjectBeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel)
		{
            if (!Validate("ProjectBeforePrint"))
            {
                Invoker.ReleaseParamsArray(pj, cancel);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("ProjectBeforePrint", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        public void ProjectBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object saveAsUi, [In] [Out] ref object cancel)
		{
            if (!Validate("ProjectBeforeSave"))
            {
                Invoker.ReleaseParamsArray(pj, saveAsUi, cancel);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
            bool newSaveAsUi = Convert.ToBoolean(saveAsUi);
			object[] paramsArray = new object[3];
			paramsArray[0] = newpj;
			paramsArray[1] = newSaveAsUi;
			paramsArray.SetValue(cancel, 2);
			EventBinding.RaiseCustomEvent("ProjectBeforeSave", ref paramsArray);

            cancel = ToBoolean(paramsArray[2]);
        }

        public void ProjectCalculate([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
		{
            if (!Validate("ProjectCalculate"))
            {
                Invoker.ReleaseParamsArray(pj);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			EventBinding.RaiseCustomEvent("ProjectCalculate", ref paramsArray);
		}

        public void WindowGoalAreaChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In] object goalArea)
		{
            if (!Validate("WindowGoalAreaChange"))
            {
                Invoker.ReleaseParamsArray(window, goalArea);
                return;
            }

			NetOffice.MSProjectApi.Window newWindow = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Window>(EventClass, window, NetOffice.MSProjectApi.Window.LateBindingApiWrapperType);
			Int32 newgoalArea = ToInt32(goalArea);
			object[] paramsArray = new object[2];
			paramsArray[0] = newWindow;
			paramsArray[1] = newgoalArea;
			EventBinding.RaiseCustomEvent("WindowGoalAreaChange", ref paramsArray);
		}

        public void WindowSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] object selType)
        {
            if (!Validate("WindowSelectionChange"))
            {
                Invoker.ReleaseParamsArray(window, sel, selType);
                return;
            }

			NetOffice.MSProjectApi.Window newWindow = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Window>(EventClass, window, NetOffice.MSProjectApi.Window.LateBindingApiWrapperType);
			NetOffice.MSProjectApi.Selection newsel = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Selection>(EventClass, sel, NetOffice.MSProjectApi.Selection.LateBindingApiWrapperType);
			object newselType = (object)selType;
			object[] paramsArray = new object[3];
			paramsArray[0] = newWindow;
			paramsArray[1] = newsel;
			paramsArray[2] = newselType;
			EventBinding.RaiseCustomEvent("WindowSelectionChange", ref paramsArray);
		}

        public void WindowBeforeViewChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In, MarshalAs(UnmanagedType.IDispatch)] object prevView, [In, MarshalAs(UnmanagedType.IDispatch)] object newView, [In] object projectHasViewWindow, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
        {
            if (!Validate("WindowBeforeViewChange"))
            {
                Invoker.ReleaseParamsArray(window, prevView, newView, projectHasViewWindow, info);
                return;
            }

			NetOffice.MSProjectApi.Window newWindow = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Window>(EventClass, window, NetOffice.MSProjectApi.Window.LateBindingApiWrapperType);
			NetOffice.MSProjectApi.View newprevView = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.View>(EventClass, prevView, NetOffice.MSProjectApi.View.LateBindingApiWrapperType);
			NetOffice.MSProjectApi.View newnewView = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.View>(EventClass, newView, NetOffice.MSProjectApi.View.LateBindingApiWrapperType);
			bool newprojectHasViewWindow = ToBoolean(projectHasViewWindow);
			NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.EventInfo>(EventClass, info, NetOffice.MSProjectApi.EventInfo.LateBindingApiWrapperType);
			object[] paramsArray = new object[5];
			paramsArray[0] = newWindow;
			paramsArray[1] = newprevView;
			paramsArray[2] = newnewView;
			paramsArray[3] = newprojectHasViewWindow;
			paramsArray[4] = newInfo;
			EventBinding.RaiseCustomEvent("WindowBeforeViewChange", ref paramsArray);
		}

        public void WindowViewChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In, MarshalAs(UnmanagedType.IDispatch)] object prevView, [In, MarshalAs(UnmanagedType.IDispatch)] object newView, [In] object success)
		{
            if (!Validate("WindowViewChange"))
            {
                Invoker.ReleaseParamsArray(window, prevView, newView, success);
                return;
            }

            NetOffice.MSProjectApi.Window newWindow = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Window>(EventClass, window, NetOffice.MSProjectApi.Window.LateBindingApiWrapperType);
            NetOffice.MSProjectApi.View newprevView = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.View>(EventClass, prevView, NetOffice.MSProjectApi.View.LateBindingApiWrapperType);
            NetOffice.MSProjectApi.View newnewView = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.View>(EventClass, newView, NetOffice.MSProjectApi.View.LateBindingApiWrapperType);
            bool newsuccess = Convert.ToBoolean(success);
			object[] paramsArray = new object[4];
			paramsArray[0] = newWindow;
			paramsArray[1] = newprevView;
			paramsArray[2] = newnewView;
			paramsArray[3] = newsuccess;
			EventBinding.RaiseCustomEvent("WindowViewChange", ref paramsArray);
		}

        public void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object activatedWindow)
		{
            if (!Validate("WindowActivate"))
            {
                Invoker.ReleaseParamsArray(activatedWindow);
                return;
            }

			NetOffice.MSProjectApi.Window newactivatedWindow = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Window>(EventClass, activatedWindow, NetOffice.MSProjectApi.Window.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newactivatedWindow;
			EventBinding.RaiseCustomEvent("WindowActivate", ref paramsArray);
		}

        public void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object deactivatedWindow)
        {
            if (!Validate("WindowDeactivate"))
            {
                Invoker.ReleaseParamsArray(deactivatedWindow);
                return;
            }

			NetOffice.MSProjectApi.Window newdeactivatedWindow = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Window>(EventClass, deactivatedWindow, NetOffice.MSProjectApi.Window.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newdeactivatedWindow;
			EventBinding.RaiseCustomEvent("WindowDeactivate", ref paramsArray);
		}

        public void WindowSidepaneDisplayChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In] object close)
		{
            if (!Validate("WindowSidepaneDisplayChange"))
            {
                Invoker.ReleaseParamsArray(window, close);
                return;
            }

            NetOffice.MSProjectApi.Window newWindow = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Window>(EventClass, window, NetOffice.MSProjectApi.Window.LateBindingApiWrapperType);
            bool newClose = Convert.ToBoolean(close);
			object[] paramsArray = new object[2];
			paramsArray[0] = newWindow;
			paramsArray[1] = newClose;
			EventBinding.RaiseCustomEvent("WindowSidepaneDisplayChange", ref paramsArray);
		}

        public void WindowSidepaneTaskChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In] object iD, [In] object isGoalArea)
        {
            if (!Validate("WindowSidepaneTaskChange"))
            {
                Invoker.ReleaseParamsArray(window, iD, isGoalArea);
                return;
            }

            NetOffice.MSProjectApi.Window newWindow = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Window>(EventClass, window, NetOffice.MSProjectApi.Window.LateBindingApiWrapperType);
            Int32 newID = ToInt32(iD);
			bool newIsGoalArea = ToBoolean(isGoalArea);
			object[] paramsArray = new object[3];
			paramsArray[0] = newWindow;
			paramsArray[1] = newID;
			paramsArray[2] = newIsGoalArea;
			EventBinding.RaiseCustomEvent("WindowSidepaneTaskChange", ref paramsArray);
		}

        public void WorkpaneDisplayChange([In] object displayState)
		{
            if (!Validate("WorkpaneDisplayChange"))
            {
                Invoker.ReleaseParamsArray(displayState);
                return;
            }

			bool newDisplayState = ToBoolean(displayState);
			object[] paramsArray = new object[1];
			paramsArray[0] = newDisplayState;
			EventBinding.RaiseCustomEvent("WorkpaneDisplayChange", ref paramsArray);
		}

        public void LoadWebPage([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In] [Out] ref object targetPage)
		{
            if (!Validate("LoadWebPage"))
            {
                Invoker.ReleaseParamsArray(window, targetPage);
                return;
            }

            NetOffice.MSProjectApi.Window newWindow = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Window>(EventClass, window, NetOffice.MSProjectApi.Window.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newWindow;
			paramsArray.SetValue(targetPage, 1);
			EventBinding.RaiseCustomEvent("LoadWebPage", ref paramsArray);

			targetPage = ToString(paramsArray[1]);
		}

		public void ProjectAfterSave()
		{
            if (!Validate("ProjectAfterSave"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("ProjectAfterSave", ref paramsArray);
		}

        public void ProjectTaskNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object iD)
		{
            if (!Validate("ProjectTaskNew"))
            {
                Invoker.ReleaseParamsArray(pj, iD);
                return;
            }

			NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
			Int32 newID = ToInt32(iD);
			object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray[1] = newID;
			EventBinding.RaiseCustomEvent("ProjectTaskNew", ref paramsArray);
		}

        public void ProjectResourceNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object iD)
        {
            if (!Validate("ProjectResourceNew"))
            {
                Invoker.ReleaseParamsArray(pj, iD);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
            Int32 newID = ToInt32(iD);
			object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray[1] = newID;
			EventBinding.RaiseCustomEvent("ProjectResourceNew", ref paramsArray);
		}

        public void ProjectAssignmentNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object iD)
		{
            if (!Validate("ProjectResourceNew"))
            {
                Invoker.ReleaseParamsArray(pj, iD);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
            Int32 newID = ToInt32(iD);
			object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray[1] = newID;
			EventBinding.RaiseCustomEvent("ProjectAssignmentNew", ref paramsArray);
		}

        public void ProjectBeforeSaveBaseline([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object interim, [In] object bl, [In] object interimCopy, [In] object interimInto, [In] object allTasks, [In] object rollupToSummaryTasks, [In] object rollupFromSubtasks, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
		{
            if (!Validate("ProjectBeforeSaveBaseline"))
            {
                Invoker.ReleaseParamsArray(pj, interim, bl, interimCopy, interimInto, allTasks, rollupToSummaryTasks, rollupFromSubtasks, info);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
            bool newInterim = Convert.ToBoolean(interim);
			NetOffice.MSProjectApi.Enums.PjBaselines newbl = (NetOffice.MSProjectApi.Enums.PjBaselines)bl;
			NetOffice.MSProjectApi.Enums.PjSaveBaselineFrom newInterimCopy = (NetOffice.MSProjectApi.Enums.PjSaveBaselineFrom)interimCopy;
			NetOffice.MSProjectApi.Enums.PjSaveBaselineTo newInterimInto = (NetOffice.MSProjectApi.Enums.PjSaveBaselineTo)interimInto;
			bool newAllTasks = ToBoolean(allTasks);
			bool newRollupToSummaryTasks = ToBoolean(rollupToSummaryTasks);
			bool newRollupFromSubtasks = ToBoolean(rollupFromSubtasks);
			NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.EventInfo>(EventClass, info, NetOffice.MSProjectApi.EventInfo.LateBindingApiWrapperType);
			object[] paramsArray = new object[9];
			paramsArray[0] = newpj;
			paramsArray[1] = newInterim;
			paramsArray[2] = newbl;
			paramsArray[3] = newInterimCopy;
			paramsArray[4] = newInterimInto;
			paramsArray[5] = newAllTasks;
			paramsArray[6] = newRollupToSummaryTasks;
			paramsArray[7] = newRollupFromSubtasks;
			paramsArray[8] = newInfo;
			EventBinding.RaiseCustomEvent("ProjectBeforeSaveBaseline", ref paramsArray);
		}

        public void ProjectBeforeClearBaseline([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object interim, [In] object bl, [In] object interimFrom, [In] object allTasks, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
		{
            if (!Validate("ProjectBeforeClearBaseline"))
            {
                Invoker.ReleaseParamsArray(pj, interim, bl, interimFrom, allTasks, info);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
            bool newInterim = ToBoolean(interim);
			NetOffice.MSProjectApi.Enums.PjBaselines newbl = (NetOffice.MSProjectApi.Enums.PjBaselines)bl;
			NetOffice.MSProjectApi.Enums.PjSaveBaselineTo newInterimFrom = (NetOffice.MSProjectApi.Enums.PjSaveBaselineTo)interimFrom;
			bool newAllTasks = ToBoolean(allTasks);
            NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.EventInfo>(EventClass, info, NetOffice.MSProjectApi.EventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[6];
			paramsArray[0] = newpj;
			paramsArray[1] = newInterim;
			paramsArray[2] = newbl;
			paramsArray[3] = newInterimFrom;
			paramsArray[4] = newAllTasks;
			paramsArray[5] = newInfo;
			EventBinding.RaiseCustomEvent("ProjectBeforeClearBaseline", ref paramsArray);
		}

        public void ProjectBeforeClose2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
        {
            if (!Validate("ProjectBeforeClose2"))
            {
                Invoker.ReleaseParamsArray(pj, info);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
            NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.EventInfo>(EventClass, info, NetOffice.MSProjectApi.EventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray[1] = newInfo;
			EventBinding.RaiseCustomEvent("ProjectBeforeClose2", ref paramsArray);
		}

        public void ProjectBeforePrint2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
        {
            if (!Validate("ProjectBeforePrint2"))
            {
                Invoker.ReleaseParamsArray(pj, info);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
            NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.EventInfo>(EventClass, info, NetOffice.MSProjectApi.EventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray[1] = newInfo;
			EventBinding.RaiseCustomEvent("ProjectBeforePrint2", ref paramsArray);
		}

        public void ProjectBeforeSave2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object saveAsUi, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
		{
            if (!Validate("ProjectBeforeSave2"))
            {
                Invoker.ReleaseParamsArray(pj, info);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
            bool newSaveAsUi = ToBoolean(saveAsUi);
            NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.EventInfo>(EventClass, info, NetOffice.MSProjectApi.EventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[3];
			paramsArray[0] = newpj;
			paramsArray[1] = newSaveAsUi;
			paramsArray[2] = newInfo;
			EventBinding.RaiseCustomEvent("ProjectBeforeSave2", ref paramsArray);
		}

        public void ProjectBeforeTaskDelete2([In, MarshalAs(UnmanagedType.IDispatch)] object tsk, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
		{
            if (!Validate("ProjectBeforeTaskDelete2"))
            {
                Invoker.ReleaseParamsArray(tsk, info);
                return;
            }

			NetOffice.MSProjectApi.Task newtsk = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Task>(EventClass, tsk, NetOffice.MSProjectApi.Task.LateBindingApiWrapperType);
            NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.EventInfo>(EventClass, info, NetOffice.MSProjectApi.EventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newtsk;
			paramsArray[1] = newInfo;
			EventBinding.RaiseCustomEvent("ProjectBeforeTaskDelete2", ref paramsArray);
		}

        public void ProjectBeforeResourceDelete2([In, MarshalAs(UnmanagedType.IDispatch)] object res, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
		{
            if (!Validate("ProjectBeforeResourceDelete2"))
            {
                Invoker.ReleaseParamsArray(res, info);
                return;
            }

			NetOffice.MSProjectApi.Resource newres = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Resource>(EventClass, res, NetOffice.MSProjectApi.Resource.LateBindingApiWrapperType);
            NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.EventInfo>(EventClass, info, NetOffice.MSProjectApi.EventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newres;
			paramsArray[1] = newInfo;
			EventBinding.RaiseCustomEvent("ProjectBeforeResourceDelete2", ref paramsArray);
		}

        public void ProjectBeforeAssignmentDelete2([In, MarshalAs(UnmanagedType.IDispatch)] object asg, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
        {
            if (!Validate("ProjectBeforeAssignmentDelete2"))
            {
                Invoker.ReleaseParamsArray(asg, info);
                return;
            }

			NetOffice.MSProjectApi.Assignment newasg = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Assignment>(EventClass, asg, NetOffice.MSProjectApi.Assignment.LateBindingApiWrapperType);
            NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.EventInfo>(EventClass, info, NetOffice.MSProjectApi.EventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newasg;
			paramsArray[1] = newInfo;
			EventBinding.RaiseCustomEvent("ProjectBeforeAssignmentDelete2", ref paramsArray);
		}

        public void ProjectBeforeTaskChange2([In, MarshalAs(UnmanagedType.IDispatch)] object tsk, [In] object field, [In] object newVal, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
        {
            if (!Validate("ProjectBeforeAssignmentDelete2"))
            {
                Invoker.ReleaseParamsArray(tsk, field, newVal, info);
                return;
            }

			NetOffice.MSProjectApi.Task newtsk = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Task>(EventClass, tsk, NetOffice.MSProjectApi.Task.LateBindingApiWrapperType);
			NetOffice.MSProjectApi.Enums.PjField newField = (NetOffice.MSProjectApi.Enums.PjField)field;
			object newNewVal = (object)newVal;
            NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.EventInfo>(EventClass, info, NetOffice.MSProjectApi.EventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[4];
			paramsArray[0] = newtsk;
			paramsArray[1] = newField;
			paramsArray[2] = newNewVal;
			paramsArray[3] = newInfo;
			EventBinding.RaiseCustomEvent("ProjectBeforeTaskChange2", ref paramsArray);
		}

        public void ProjectBeforeResourceChange2([In, MarshalAs(UnmanagedType.IDispatch)] object res, [In] object field, [In] object newVal, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
		{
            if (!Validate("ProjectBeforeResourceChange2"))
            {
                Invoker.ReleaseParamsArray(res, field, newVal, info);
                return;
            }

			NetOffice.MSProjectApi.Resource newres = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Resource>(EventClass, res, NetOffice.MSProjectApi.Resource.LateBindingApiWrapperType);
			NetOffice.MSProjectApi.Enums.PjField newField = (NetOffice.MSProjectApi.Enums.PjField)field;
			object newNewVal = (object)newVal;
            NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.EventInfo>(EventClass, info, NetOffice.MSProjectApi.EventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[4];
			paramsArray[0] = newres;
			paramsArray[1] = newField;
			paramsArray[2] = newNewVal;
			paramsArray[3] = newInfo;
			EventBinding.RaiseCustomEvent("ProjectBeforeResourceChange2", ref paramsArray);
		}

        public void ProjectBeforeAssignmentChange2([In, MarshalAs(UnmanagedType.IDispatch)] object asg, [In] object field, [In] object newVal, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
		{
            if (!Validate("ProjectBeforeAssignmentChange2"))
            {
                Invoker.ReleaseParamsArray(asg, field, newVal, info);
                return;
            }

			NetOffice.MSProjectApi.Assignment newasg = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Assignment>(EventClass, asg, NetOffice.MSProjectApi.Assignment.LateBindingApiWrapperType);
			NetOffice.MSProjectApi.Enums.PjAssignmentField newField = (NetOffice.MSProjectApi.Enums.PjAssignmentField)field;
			object newNewVal = (object)newVal;
            NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.EventInfo>(EventClass, info, NetOffice.MSProjectApi.EventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[4];
			paramsArray[0] = newasg;
			paramsArray[1] = newField;
			paramsArray[2] = newNewVal;
			paramsArray[3] = newInfo;
			EventBinding.RaiseCustomEvent("ProjectBeforeAssignmentChange2", ref paramsArray);
		}

        public void ProjectBeforeTaskNew2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
        {
            if (!Validate("ProjectBeforeTaskNew2"))
            {
                Invoker.ReleaseParamsArray(pj, info);
                return;
            }

			NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
            NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.EventInfo>(EventClass, info, NetOffice.MSProjectApi.EventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray[1] = newInfo;
			EventBinding.RaiseCustomEvent("ProjectBeforeTaskNew2", ref paramsArray);
		}

        public void ProjectBeforeResourceNew2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
        {
            if (!Validate("ProjectBeforeResourceNew2"))
            {
                Invoker.ReleaseParamsArray(pj, info);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
            NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.EventInfo>(EventClass, info, NetOffice.MSProjectApi.EventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray[1] = newInfo;
			EventBinding.RaiseCustomEvent("ProjectBeforeResourceNew2", ref paramsArray);
		}

        public void ProjectBeforeAssignmentNew2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
		{
            if (!Validate("ProjectBeforeResourceNew2"))
            {
                Invoker.ReleaseParamsArray(pj, info);
                return;
            }

            NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
            NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.EventInfo>(EventClass, info, NetOffice.MSProjectApi.EventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray[1] = newInfo;
			EventBinding.RaiseCustomEvent("ProjectBeforeAssignmentNew2", ref paramsArray);
		}

        public void ApplicationBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object info)
		{
            if (!Validate("ApplicationBeforeClose"))
            {
                Invoker.ReleaseParamsArray(info);
                return;
            }

            NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.EventInfo>(EventClass, info, NetOffice.MSProjectApi.EventInfo.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newInfo;
			EventBinding.RaiseCustomEvent("ApplicationBeforeClose", ref paramsArray);
		}

        public void OnUndoOrRedo([In] object bstrLabel, [In] object bstrGUID, [In] object fUndo)
        {
            if (!Validate("OnUndoOrRedo"))
            {
                Invoker.ReleaseParamsArray(bstrLabel, bstrGUID, fUndo);
                return;
            }

			string newbstrLabel = ToString(bstrLabel);
			string newbstrGUID = ToString(bstrGUID);
			bool newfUndo = ToBoolean(fUndo);
			object[] paramsArray = new object[3];
			paramsArray[0] = newbstrLabel;
			paramsArray[1] = newbstrGUID;
			paramsArray[2] = newfUndo;
			EventBinding.RaiseCustomEvent("OnUndoOrRedo", ref paramsArray);
		}

        public void AfterCubeBuilt([In] [Out] ref object cubeFileName)
        {
            if (!Validate("AfterCubeBuilt"))
            {
                Invoker.ReleaseParamsArray(cubeFileName);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cubeFileName, 0);
			EventBinding.RaiseCustomEvent("AfterCubeBuilt", ref paramsArray);

			cubeFileName = ToString(paramsArray[0]);
		}

        public void LoadWebPane([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In] [Out] ref object targetPage)
		{
            if(!Validate("LoadWebPane"))
            {
                Invoker.ReleaseParamsArray(window, targetPage);
                return;
            }

			NetOffice.MSProjectApi.Window newWindow = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Window>(EventClass, window, NetOffice.MSProjectApi.Window.LateBindingApiWrapperType);
			object[] paramsArray = new object[2];
			paramsArray[0] = newWindow;
			paramsArray.SetValue(targetPage, 1);
			EventBinding.RaiseCustomEvent("LoadWebPane", ref paramsArray);

			targetPage = ToString(paramsArray[1]);
        }

        public void JobStart([In] object bstrName, [In] object bstrprojGuid, [In] object bstrjobGuid, [In] object jobType, [In] object lResult)
		{
            if (!Validate("JobStart"))
            {
                Invoker.ReleaseParamsArray(bstrName, bstrprojGuid, bstrjobGuid, jobType, lResult);
                return;
            }

			string newbstrName = ToString(bstrName);
			string newbstrprojGuid = ToString(bstrprojGuid);
			string newbstrjobGuid = ToString(bstrjobGuid);
			Int32 newjobType = ToInt32(jobType);
			Int32 newlResult = ToInt32(lResult);
			object[] paramsArray = new object[5];
			paramsArray[0] = newbstrName;
			paramsArray[1] = newbstrprojGuid;
			paramsArray[2] = newbstrjobGuid;
			paramsArray[3] = newjobType;
			paramsArray[4] = newlResult;
			EventBinding.RaiseCustomEvent("JobStart", ref paramsArray);
		}

        public void JobCompleted([In] object bstrName, [In] object bstrprojGuid, [In] object bstrjobGuid, [In] object jobType, [In] object lResult)
		{
            if (!Validate("JobCompleted"))
            {
                Invoker.ReleaseParamsArray(bstrName, bstrprojGuid, bstrjobGuid, jobType, lResult);
                return;
            }

			string newbstrName = ToString(bstrName);
			string newbstrprojGuid = ToString(bstrprojGuid);
			string newbstrjobGuid = ToString(bstrjobGuid);
			Int32 newjobType = ToInt32(jobType);
			Int32 newlResult = ToInt32(lResult);
			object[] paramsArray = new object[5];
			paramsArray[0] = newbstrName;
			paramsArray[1] = newbstrprojGuid;
			paramsArray[2] = newbstrjobGuid;
			paramsArray[3] = newjobType;
			paramsArray[4] = newlResult;
			EventBinding.RaiseCustomEvent("JobCompleted", ref paramsArray);
		}

        public void SaveStartingToServer([In] object bstrName, [In] object bstrprojGuid)
		{
            if (!Validate("SaveStartingToServer"))
            {
                Invoker.ReleaseParamsArray(bstrName, bstrprojGuid);
                return;
            }

			string newbstrName = ToString(bstrName);
			string newbstrprojGuid = ToString(bstrprojGuid);
			object[] paramsArray = new object[2];
			paramsArray[0] = newbstrName;
			paramsArray[1] = newbstrprojGuid;
			EventBinding.RaiseCustomEvent("SaveStartingToServer", ref paramsArray);
		}

        public void SaveCompletedToServer([In] object bstrName, [In] object bstrprojGuid)
		{
            if (!Validate("SaveCompletedToServer"))
            {
                Invoker.ReleaseParamsArray(bstrName, bstrprojGuid);
                return;
            }

			string newbstrName = ToString(bstrName);
			string newbstrprojGuid = ToString(bstrprojGuid);
			object[] paramsArray = new object[2];
			paramsArray[0] = newbstrName;
			paramsArray[1] = newbstrprojGuid;
			EventBinding.RaiseCustomEvent("SaveCompletedToServer", ref paramsArray);
		}

        public void ProjectBeforePublish([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel)
        {
            if (!Validate("ProjectBeforePublish"))
            {
                Invoker.ReleaseParamsArray(pj, cancel);
                return;
            }

			NetOffice.MSProjectApi.Project newpj = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Project>(EventClass, pj, NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
			object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("ProjectBeforePublish", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
		}

		public void PaneActivate()
		{
            if (!Validate("PaneActivate"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("PaneActivate", ref paramsArray);
		}

        public void SecondaryViewChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In, MarshalAs(UnmanagedType.IDispatch)] object prevView, [In, MarshalAs(UnmanagedType.IDispatch)] object newView, [In] object success)
		{
            if (!Validate("SecondaryViewChange"))
            {
                Invoker.ReleaseParamsArray(window, prevView, newView, success);
                return;
            }

			NetOffice.MSProjectApi.Window newWindow = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.Window>(EventClass, window, NetOffice.MSProjectApi.Window.LateBindingApiWrapperType);
			NetOffice.MSProjectApi.View newprevView = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.View>(EventClass, prevView, NetOffice.MSProjectApi.View.LateBindingApiWrapperType);
            NetOffice.MSProjectApi.View newnewView = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.View> (EventClass, newView, NetOffice.MSProjectApi.View.LateBindingApiWrapperType);
			bool newsuccess = ToBoolean(success);
			object[] paramsArray = new object[4];
			paramsArray[0] = newWindow;
			paramsArray[1] = newprevView;
			paramsArray[2] = newnewView;
			paramsArray[3] = newsuccess;
			EventBinding.RaiseCustomEvent("SecondaryViewChange", ref paramsArray);
		}

        public void IsFunctionalitySupported([In] object bstrFunctionality, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
        {
            if (!Validate("IsFunctionalitySupported"))
            {
                Invoker.ReleaseParamsArray(bstrFunctionality, info);
                return;
            }

			string newbstrFunctionality = ToString(bstrFunctionality);
			NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateKnownObjectFromComProxy<NetOffice.MSProjectApi.EventInfo>(EventClass, info, NetOffice.MSProjectApi.EventInfo.LateBindingApiWrapperType);
			object[] paramsArray = new object[2];
			paramsArray[0] = newbstrFunctionality;
			paramsArray[1] = newInfo;
			EventBinding.RaiseCustomEvent("IsFunctionalitySupported", ref paramsArray);
		}

        public void ConnectionStatusChanged([In] object online)
        {
            if (!Validate("IsFunctionalitySupported"))
            {
                Invoker.ReleaseParamsArray(online);
                return;
            }

			bool newonline = ToBoolean(online);
			object[] paramsArray = new object[1];
			paramsArray[0] = newonline;
			EventBinding.RaiseCustomEvent("ConnectionStatusChanged", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}