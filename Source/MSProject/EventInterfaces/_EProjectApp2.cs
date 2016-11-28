using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;

namespace NetOffice.MSProjectApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersionAttribute("MSProject", 11,12,14)]
	[ComImport, Guid("5066D7C4-1ED7-48C4-ACE7-299E109D368C"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _EProjectApp2
	{
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void NewProject([In, MarshalAs(UnmanagedType.IDispatch)] object pj);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
		void ProjectBeforeTaskDelete([In, MarshalAs(UnmanagedType.IDispatch)] object tsk, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)]
		void ProjectBeforeResourceDelete([In, MarshalAs(UnmanagedType.IDispatch)] object res, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)]
		void ProjectBeforeAssignmentDelete([In, MarshalAs(UnmanagedType.IDispatch)] object asg, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(9)]
		void ProjectBeforeTaskChange([In, MarshalAs(UnmanagedType.IDispatch)] object tsk, [In] object field, [In] object newVal, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(10)]
		void ProjectBeforeResourceChange([In, MarshalAs(UnmanagedType.IDispatch)] object res, [In] object field, [In] object newVal, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(11)]
		void ProjectBeforeAssignmentChange([In, MarshalAs(UnmanagedType.IDispatch)] object asg, [In] object field, [In] object newVal, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(12)]
		void ProjectBeforeTaskNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(13)]
		void ProjectBeforeResourceNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(14)]
		void ProjectBeforeAssignmentNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void ProjectBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void ProjectBeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void ProjectBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object saveAsUi, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)]
		void ProjectCalculate([In, MarshalAs(UnmanagedType.IDispatch)] object pj);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(15)]
		void WindowGoalAreaChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In] object goalArea);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16)]
		void WindowSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] object selType);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(17)]
		void WindowBeforeViewChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In, MarshalAs(UnmanagedType.IDispatch)] object prevView, [In, MarshalAs(UnmanagedType.IDispatch)] object newView, [In] object projectHasViewWindow, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(18)]
		void WindowViewChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In, MarshalAs(UnmanagedType.IDispatch)] object prevView, [In, MarshalAs(UnmanagedType.IDispatch)] object newView, [In] object success);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(19)]
		void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object activatedWindow);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(20)]
		void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object deactivatedWindow);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(21)]
		void WindowSidepaneDisplayChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In] object close);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(22)]
		void WindowSidepaneTaskChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In] object iD, [In] object isGoalArea);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(23)]
		void WorkpaneDisplayChange([In] object displayState);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(24)]
		void LoadWebPage([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In] [Out] ref object targetPage);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(25)]
		void ProjectAfterSave();

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(26)]
		void ProjectTaskNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object iD);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(27)]
		void ProjectResourceNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object iD);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(28)]
		void ProjectAssignmentNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object iD);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(29)]
		void ProjectBeforeSaveBaseline([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object interim, [In] object bl, [In] object interimCopy, [In] object interimInto, [In] object allTasks, [In] object rollupToSummaryTasks, [In] object rollupFromSubtasks, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(30)]
		void ProjectBeforeClearBaseline([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object interim, [In] object bl, [In] object interimFrom, [In] object allTasks, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741826)]
		void ProjectBeforeClose2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741828)]
		void ProjectBeforePrint2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741827)]
		void ProjectBeforeSave2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object saveAsUi, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741830)]
		void ProjectBeforeTaskDelete2([In, MarshalAs(UnmanagedType.IDispatch)] object tsk, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741831)]
		void ProjectBeforeResourceDelete2([In, MarshalAs(UnmanagedType.IDispatch)] object res, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741832)]
		void ProjectBeforeAssignmentDelete2([In, MarshalAs(UnmanagedType.IDispatch)] object asg, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741833)]
		void ProjectBeforeTaskChange2([In, MarshalAs(UnmanagedType.IDispatch)] object tsk, [In] object field, [In] object newVal, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741834)]
		void ProjectBeforeResourceChange2([In, MarshalAs(UnmanagedType.IDispatch)] object res, [In] object field, [In] object newVal, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741835)]
		void ProjectBeforeAssignmentChange2([In, MarshalAs(UnmanagedType.IDispatch)] object asg, [In] object field, [In] object newVal, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741836)]
		void ProjectBeforeTaskNew2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741837)]
		void ProjectBeforeResourceNew2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1073741838)]
		void ProjectBeforeAssignmentNew2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(31)]
		void ApplicationBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(32)]
		void OnUndoOrRedo([In] object bstrLabel, [In] object bstrGUID, [In] object fUndo);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(33)]
		void AfterCubeBuilt([In] [Out] ref object cubeFileName);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(34)]
		void LoadWebPane([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In] [Out] ref object targetPage);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(35)]
		void JobStart([In] object bstrName, [In] object bstrprojGuid, [In] object bstrjobGuid, [In] object jobType, [In] object lResult);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(36)]
		void JobCompleted([In] object bstrName, [In] object bstrprojGuid, [In] object bstrjobGuid, [In] object jobType, [In] object lResult);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(37)]
		void SaveStartingToServer([In] object bstrName, [In] object bstrprojGuid);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(38)]
		void SaveCompletedToServer([In] object bstrName, [In] object bstrprojGuid);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(39)]
		void ProjectBeforePublish([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(40)]
		void PaneActivate();

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(41)]
		void SecondaryViewChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In, MarshalAs(UnmanagedType.IDispatch)] object prevView, [In, MarshalAs(UnmanagedType.IDispatch)] object newView, [In] object success);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(42)]
		void IsFunctionalitySupported([In] object bstrFunctionality, [In, MarshalAs(UnmanagedType.IDispatch)] object info);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(43)]
		void ConnectionStatusChanged([In] object online);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _EProjectApp2_SinkHelper : SinkHelper, _EProjectApp2
	{
		#region Static
		
		public static readonly string Id = "5066D7C4-1ED7-48C4-ACE7-299E109D368C";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public _EProjectApp2_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region Properties

        internal Core Factory
        {
            get
            {
                if (null != _eventClass)
                    return _eventClass.Factory;
                else
                    return Core.Default;
            }
        }

        #endregion

		#region _EProjectApp2 Members
		
		public void NewProject([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("NewProject");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			_eventBinding.RaiseCustomEvent("NewProject", ref paramsArray);
		}

		public void ProjectBeforeTaskDelete([In, MarshalAs(UnmanagedType.IDispatch)] object tsk, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectBeforeTaskDelete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(tsk, cancel);
				return;
			}

			NetOffice.MSProjectApi.Task newtsk = Factory.CreateObjectFromComProxy(_eventClass, tsk) as NetOffice.MSProjectApi.Task;
			object[] paramsArray = new object[2];
			paramsArray[0] = newtsk;
			paramsArray.SetValue(cancel, 1);
			_eventBinding.RaiseCustomEvent("ProjectBeforeTaskDelete", ref paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void ProjectBeforeResourceDelete([In, MarshalAs(UnmanagedType.IDispatch)] object res, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectBeforeResourceDelete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(res, cancel);
				return;
			}

			NetOffice.MSProjectApi.Resource newres = Factory.CreateObjectFromComProxy(_eventClass, res) as NetOffice.MSProjectApi.Resource;
			object[] paramsArray = new object[2];
			paramsArray[0] = newres;
			paramsArray.SetValue(cancel, 1);
			_eventBinding.RaiseCustomEvent("ProjectBeforeResourceDelete", ref paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void ProjectBeforeAssignmentDelete([In, MarshalAs(UnmanagedType.IDispatch)] object asg, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectBeforeAssignmentDelete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(asg, cancel);
				return;
			}

			NetOffice.MSProjectApi.Assignment newasg = Factory.CreateObjectFromComProxy(_eventClass, asg) as NetOffice.MSProjectApi.Assignment;
			object[] paramsArray = new object[2];
			paramsArray[0] = newasg;
			paramsArray.SetValue(cancel, 1);
			_eventBinding.RaiseCustomEvent("ProjectBeforeAssignmentDelete", ref paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void ProjectBeforeTaskChange([In, MarshalAs(UnmanagedType.IDispatch)] object tsk, [In] object field, [In] object newVal, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectBeforeTaskChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(tsk, field, newVal, cancel);
				return;
			}

			NetOffice.MSProjectApi.Task newtsk = Factory.CreateObjectFromComProxy(_eventClass, tsk) as NetOffice.MSProjectApi.Task;
			NetOffice.MSProjectApi.Enums.PjField newField = (NetOffice.MSProjectApi.Enums.PjField)field;
			object newNewVal = (object)newVal;
			object[] paramsArray = new object[4];
			paramsArray[0] = newtsk;
			paramsArray[1] = newField;
			paramsArray[2] = newNewVal;
			paramsArray.SetValue(cancel, 3);
			_eventBinding.RaiseCustomEvent("ProjectBeforeTaskChange", ref paramsArray);

			cancel = (bool)paramsArray[3];
		}

		public void ProjectBeforeResourceChange([In, MarshalAs(UnmanagedType.IDispatch)] object res, [In] object field, [In] object newVal, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectBeforeResourceChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(res, field, newVal, cancel);
				return;
			}

			NetOffice.MSProjectApi.Resource newres = Factory.CreateObjectFromComProxy(_eventClass, res) as NetOffice.MSProjectApi.Resource;
			NetOffice.MSProjectApi.Enums.PjField newField = (NetOffice.MSProjectApi.Enums.PjField)field;
			object newNewVal = (object)newVal;
			object[] paramsArray = new object[4];
			paramsArray[0] = newres;
			paramsArray[1] = newField;
			paramsArray[2] = newNewVal;
			paramsArray.SetValue(cancel, 3);
			_eventBinding.RaiseCustomEvent("ProjectBeforeResourceChange", ref paramsArray);

			cancel = (bool)paramsArray[3];
		}

		public void ProjectBeforeAssignmentChange([In, MarshalAs(UnmanagedType.IDispatch)] object asg, [In] object field, [In] object newVal, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectBeforeAssignmentChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(asg, field, newVal, cancel);
				return;
			}

			NetOffice.MSProjectApi.Assignment newasg = Factory.CreateObjectFromComProxy(_eventClass, asg) as NetOffice.MSProjectApi.Assignment;
			NetOffice.MSProjectApi.Enums.PjAssignmentField newField = (NetOffice.MSProjectApi.Enums.PjAssignmentField)field;
			object newNewVal = (object)newVal;
			object[] paramsArray = new object[4];
			paramsArray[0] = newasg;
			paramsArray[1] = newField;
			paramsArray[2] = newNewVal;
			paramsArray.SetValue(cancel, 3);
			_eventBinding.RaiseCustomEvent("ProjectBeforeAssignmentChange", ref paramsArray);

			cancel = (bool)paramsArray[3];
		}

		public void ProjectBeforeTaskNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectBeforeTaskNew");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj, cancel);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray.SetValue(cancel, 1);
			_eventBinding.RaiseCustomEvent("ProjectBeforeTaskNew", ref paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void ProjectBeforeResourceNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectBeforeResourceNew");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj, cancel);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray.SetValue(cancel, 1);
			_eventBinding.RaiseCustomEvent("ProjectBeforeResourceNew", ref paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void ProjectBeforeAssignmentNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectBeforeAssignmentNew");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj, cancel);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray.SetValue(cancel, 1);
			_eventBinding.RaiseCustomEvent("ProjectBeforeAssignmentNew", ref paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void ProjectBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectBeforeClose");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj, cancel);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray.SetValue(cancel, 1);
			_eventBinding.RaiseCustomEvent("ProjectBeforeClose", ref paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void ProjectBeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectBeforePrint");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj, cancel);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray.SetValue(cancel, 1);
			_eventBinding.RaiseCustomEvent("ProjectBeforePrint", ref paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void ProjectBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object saveAsUi, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectBeforeSave");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj, saveAsUi, cancel);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			bool newSaveAsUi = Convert.ToBoolean(saveAsUi);
			object[] paramsArray = new object[3];
			paramsArray[0] = newpj;
			paramsArray[1] = newSaveAsUi;
			paramsArray.SetValue(cancel, 2);
			_eventBinding.RaiseCustomEvent("ProjectBeforeSave", ref paramsArray);

			cancel = (bool)paramsArray[2];
		}

		public void ProjectCalculate([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectCalculate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			_eventBinding.RaiseCustomEvent("ProjectCalculate", ref paramsArray);
		}

		public void WindowGoalAreaChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In] object goalArea)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowGoalAreaChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window, goalArea);
				return;
			}

			NetOffice.MSProjectApi.Window newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.MSProjectApi.Window;
			Int32 newgoalArea = Convert.ToInt32(goalArea);
			object[] paramsArray = new object[2];
			paramsArray[0] = newWindow;
			paramsArray[1] = newgoalArea;
			_eventBinding.RaiseCustomEvent("WindowGoalAreaChange", ref paramsArray);
		}

		public void WindowSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] object selType)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowSelectionChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window, sel, selType);
				return;
			}

			NetOffice.MSProjectApi.Window newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.MSProjectApi.Window;
			NetOffice.MSProjectApi.Selection newsel = Factory.CreateObjectFromComProxy(_eventClass, sel) as NetOffice.MSProjectApi.Selection;
			object newselType = (object)selType;
			object[] paramsArray = new object[3];
			paramsArray[0] = newWindow;
			paramsArray[1] = newsel;
			paramsArray[2] = newselType;
			_eventBinding.RaiseCustomEvent("WindowSelectionChange", ref paramsArray);
		}

		public void WindowBeforeViewChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In, MarshalAs(UnmanagedType.IDispatch)] object prevView, [In, MarshalAs(UnmanagedType.IDispatch)] object newView, [In] object projectHasViewWindow, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowBeforeViewChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window, prevView, newView, projectHasViewWindow, info);
				return;
			}

			NetOffice.MSProjectApi.Window newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.MSProjectApi.Window;
			NetOffice.MSProjectApi.View newprevView = Factory.CreateObjectFromComProxy(_eventClass, prevView) as NetOffice.MSProjectApi.View;
			NetOffice.MSProjectApi.View newnewView = Factory.CreateObjectFromComProxy(_eventClass, newView) as NetOffice.MSProjectApi.View;
			bool newprojectHasViewWindow = Convert.ToBoolean(projectHasViewWindow);
			NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateObjectFromComProxy(_eventClass, info) as NetOffice.MSProjectApi.EventInfo;
			object[] paramsArray = new object[5];
			paramsArray[0] = newWindow;
			paramsArray[1] = newprevView;
			paramsArray[2] = newnewView;
			paramsArray[3] = newprojectHasViewWindow;
			paramsArray[4] = newInfo;
			_eventBinding.RaiseCustomEvent("WindowBeforeViewChange", ref paramsArray);
		}

		public void WindowViewChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In, MarshalAs(UnmanagedType.IDispatch)] object prevView, [In, MarshalAs(UnmanagedType.IDispatch)] object newView, [In] object success)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowViewChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window, prevView, newView, success);
				return;
			}

			NetOffice.MSProjectApi.Window newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.MSProjectApi.Window;
			NetOffice.MSProjectApi.View newprevView = Factory.CreateObjectFromComProxy(_eventClass, prevView) as NetOffice.MSProjectApi.View;
			NetOffice.MSProjectApi.View newnewView = Factory.CreateObjectFromComProxy(_eventClass, newView) as NetOffice.MSProjectApi.View;
			bool newsuccess = Convert.ToBoolean(success);
			object[] paramsArray = new object[4];
			paramsArray[0] = newWindow;
			paramsArray[1] = newprevView;
			paramsArray[2] = newnewView;
			paramsArray[3] = newsuccess;
			_eventBinding.RaiseCustomEvent("WindowViewChange", ref paramsArray);
		}

		public void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object activatedWindow)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowActivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(activatedWindow);
				return;
			}

			NetOffice.MSProjectApi.Window newactivatedWindow = Factory.CreateObjectFromComProxy(_eventClass, activatedWindow) as NetOffice.MSProjectApi.Window;
			object[] paramsArray = new object[1];
			paramsArray[0] = newactivatedWindow;
			_eventBinding.RaiseCustomEvent("WindowActivate", ref paramsArray);
		}

		public void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object deactivatedWindow)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowDeactivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(deactivatedWindow);
				return;
			}

			NetOffice.MSProjectApi.Window newdeactivatedWindow = Factory.CreateObjectFromComProxy(_eventClass, deactivatedWindow) as NetOffice.MSProjectApi.Window;
			object[] paramsArray = new object[1];
			paramsArray[0] = newdeactivatedWindow;
			_eventBinding.RaiseCustomEvent("WindowDeactivate", ref paramsArray);
		}

		public void WindowSidepaneDisplayChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In] object close)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowSidepaneDisplayChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window, close);
				return;
			}

			NetOffice.MSProjectApi.Window newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.MSProjectApi.Window;
			bool newClose = Convert.ToBoolean(close);
			object[] paramsArray = new object[2];
			paramsArray[0] = newWindow;
			paramsArray[1] = newClose;
			_eventBinding.RaiseCustomEvent("WindowSidepaneDisplayChange", ref paramsArray);
		}

		public void WindowSidepaneTaskChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In] object iD, [In] object isGoalArea)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowSidepaneTaskChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window, iD, isGoalArea);
				return;
			}

			NetOffice.MSProjectApi.Window newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.MSProjectApi.Window;
			Int32 newID = Convert.ToInt32(iD);
			bool newIsGoalArea = Convert.ToBoolean(isGoalArea);
			object[] paramsArray = new object[3];
			paramsArray[0] = newWindow;
			paramsArray[1] = newID;
			paramsArray[2] = newIsGoalArea;
			_eventBinding.RaiseCustomEvent("WindowSidepaneTaskChange", ref paramsArray);
		}

		public void WorkpaneDisplayChange([In] object displayState)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WorkpaneDisplayChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(displayState);
				return;
			}

			bool newDisplayState = Convert.ToBoolean(displayState);
			object[] paramsArray = new object[1];
			paramsArray[0] = newDisplayState;
			_eventBinding.RaiseCustomEvent("WorkpaneDisplayChange", ref paramsArray);
		}

		public void LoadWebPage([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In] [Out] ref object targetPage)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("LoadWebPage");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window, targetPage);
				return;
			}

			NetOffice.MSProjectApi.Window newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.MSProjectApi.Window;
			object[] paramsArray = new object[2];
			paramsArray[0] = newWindow;
			paramsArray.SetValue(targetPage, 1);
			_eventBinding.RaiseCustomEvent("LoadWebPage", ref paramsArray);

			targetPage = (string)paramsArray[1];
		}

		public void ProjectAfterSave()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectAfterSave");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("ProjectAfterSave", ref paramsArray);
		}

		public void ProjectTaskNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object iD)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectTaskNew");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj, iD);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			Int32 newID = Convert.ToInt32(iD);
			object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray[1] = newID;
			_eventBinding.RaiseCustomEvent("ProjectTaskNew", ref paramsArray);
		}

		public void ProjectResourceNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object iD)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectResourceNew");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj, iD);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			Int32 newID = Convert.ToInt32(iD);
			object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray[1] = newID;
			_eventBinding.RaiseCustomEvent("ProjectResourceNew", ref paramsArray);
		}

		public void ProjectAssignmentNew([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object iD)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectAssignmentNew");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj, iD);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			Int32 newID = Convert.ToInt32(iD);
			object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray[1] = newID;
			_eventBinding.RaiseCustomEvent("ProjectAssignmentNew", ref paramsArray);
		}

		public void ProjectBeforeSaveBaseline([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object interim, [In] object bl, [In] object interimCopy, [In] object interimInto, [In] object allTasks, [In] object rollupToSummaryTasks, [In] object rollupFromSubtasks, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectBeforeSaveBaseline");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj, interim, bl, interimCopy, interimInto, allTasks, rollupToSummaryTasks, rollupFromSubtasks, info);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			bool newInterim = Convert.ToBoolean(interim);
			NetOffice.MSProjectApi.Enums.PjBaselines newbl = (NetOffice.MSProjectApi.Enums.PjBaselines)bl;
			NetOffice.MSProjectApi.Enums.PjSaveBaselineFrom newInterimCopy = (NetOffice.MSProjectApi.Enums.PjSaveBaselineFrom)interimCopy;
			NetOffice.MSProjectApi.Enums.PjSaveBaselineTo newInterimInto = (NetOffice.MSProjectApi.Enums.PjSaveBaselineTo)interimInto;
			bool newAllTasks = Convert.ToBoolean(allTasks);
			bool newRollupToSummaryTasks = Convert.ToBoolean(rollupToSummaryTasks);
			bool newRollupFromSubtasks = Convert.ToBoolean(rollupFromSubtasks);
			NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateObjectFromComProxy(_eventClass, info) as NetOffice.MSProjectApi.EventInfo;
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
			_eventBinding.RaiseCustomEvent("ProjectBeforeSaveBaseline", ref paramsArray);
		}

		public void ProjectBeforeClearBaseline([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object interim, [In] object bl, [In] object interimFrom, [In] object allTasks, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectBeforeClearBaseline");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj, interim, bl, interimFrom, allTasks, info);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			bool newInterim = Convert.ToBoolean(interim);
			NetOffice.MSProjectApi.Enums.PjBaselines newbl = (NetOffice.MSProjectApi.Enums.PjBaselines)bl;
			NetOffice.MSProjectApi.Enums.PjSaveBaselineTo newInterimFrom = (NetOffice.MSProjectApi.Enums.PjSaveBaselineTo)interimFrom;
			bool newAllTasks = Convert.ToBoolean(allTasks);
			NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateObjectFromComProxy(_eventClass, info) as NetOffice.MSProjectApi.EventInfo;
			object[] paramsArray = new object[6];
			paramsArray[0] = newpj;
			paramsArray[1] = newInterim;
			paramsArray[2] = newbl;
			paramsArray[3] = newInterimFrom;
			paramsArray[4] = newAllTasks;
			paramsArray[5] = newInfo;
			_eventBinding.RaiseCustomEvent("ProjectBeforeClearBaseline", ref paramsArray);
		}

		public void ProjectBeforeClose2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectBeforeClose2");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj, info);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateObjectFromComProxy(_eventClass, info) as NetOffice.MSProjectApi.EventInfo;
			object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray[1] = newInfo;
			_eventBinding.RaiseCustomEvent("ProjectBeforeClose2", ref paramsArray);
		}

		public void ProjectBeforePrint2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectBeforePrint2");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj, info);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateObjectFromComProxy(_eventClass, info) as NetOffice.MSProjectApi.EventInfo;
			object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray[1] = newInfo;
			_eventBinding.RaiseCustomEvent("ProjectBeforePrint2", ref paramsArray);
		}

		public void ProjectBeforeSave2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] object saveAsUi, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectBeforeSave2");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj, saveAsUi, info);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			bool newSaveAsUi = Convert.ToBoolean(saveAsUi);
			NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateObjectFromComProxy(_eventClass, info) as NetOffice.MSProjectApi.EventInfo;
			object[] paramsArray = new object[3];
			paramsArray[0] = newpj;
			paramsArray[1] = newSaveAsUi;
			paramsArray[2] = newInfo;
			_eventBinding.RaiseCustomEvent("ProjectBeforeSave2", ref paramsArray);
		}

		public void ProjectBeforeTaskDelete2([In, MarshalAs(UnmanagedType.IDispatch)] object tsk, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectBeforeTaskDelete2");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(tsk, info);
				return;
			}

			NetOffice.MSProjectApi.Task newtsk = Factory.CreateObjectFromComProxy(_eventClass, tsk) as NetOffice.MSProjectApi.Task;
			NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateObjectFromComProxy(_eventClass, info) as NetOffice.MSProjectApi.EventInfo;
			object[] paramsArray = new object[2];
			paramsArray[0] = newtsk;
			paramsArray[1] = newInfo;
			_eventBinding.RaiseCustomEvent("ProjectBeforeTaskDelete2", ref paramsArray);
		}

		public void ProjectBeforeResourceDelete2([In, MarshalAs(UnmanagedType.IDispatch)] object res, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectBeforeResourceDelete2");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(res, info);
				return;
			}

			NetOffice.MSProjectApi.Resource newres = Factory.CreateObjectFromComProxy(_eventClass, res) as NetOffice.MSProjectApi.Resource;
			NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateObjectFromComProxy(_eventClass, info) as NetOffice.MSProjectApi.EventInfo;
			object[] paramsArray = new object[2];
			paramsArray[0] = newres;
			paramsArray[1] = newInfo;
			_eventBinding.RaiseCustomEvent("ProjectBeforeResourceDelete2", ref paramsArray);
		}

		public void ProjectBeforeAssignmentDelete2([In, MarshalAs(UnmanagedType.IDispatch)] object asg, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectBeforeAssignmentDelete2");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(asg, info);
				return;
			}

			NetOffice.MSProjectApi.Assignment newasg = Factory.CreateObjectFromComProxy(_eventClass, asg) as NetOffice.MSProjectApi.Assignment;
			NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateObjectFromComProxy(_eventClass, info) as NetOffice.MSProjectApi.EventInfo;
			object[] paramsArray = new object[2];
			paramsArray[0] = newasg;
			paramsArray[1] = newInfo;
			_eventBinding.RaiseCustomEvent("ProjectBeforeAssignmentDelete2", ref paramsArray);
		}

		public void ProjectBeforeTaskChange2([In, MarshalAs(UnmanagedType.IDispatch)] object tsk, [In] object field, [In] object newVal, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectBeforeTaskChange2");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(tsk, field, newVal, info);
				return;
			}

			NetOffice.MSProjectApi.Task newtsk = Factory.CreateObjectFromComProxy(_eventClass, tsk) as NetOffice.MSProjectApi.Task;
			NetOffice.MSProjectApi.Enums.PjField newField = (NetOffice.MSProjectApi.Enums.PjField)field;
			object newNewVal = (object)newVal;
			NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateObjectFromComProxy(_eventClass, info) as NetOffice.MSProjectApi.EventInfo;
			object[] paramsArray = new object[4];
			paramsArray[0] = newtsk;
			paramsArray[1] = newField;
			paramsArray[2] = newNewVal;
			paramsArray[3] = newInfo;
			_eventBinding.RaiseCustomEvent("ProjectBeforeTaskChange2", ref paramsArray);
		}

		public void ProjectBeforeResourceChange2([In, MarshalAs(UnmanagedType.IDispatch)] object res, [In] object field, [In] object newVal, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectBeforeResourceChange2");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(res, field, newVal, info);
				return;
			}

			NetOffice.MSProjectApi.Resource newres = Factory.CreateObjectFromComProxy(_eventClass, res) as NetOffice.MSProjectApi.Resource;
			NetOffice.MSProjectApi.Enums.PjField newField = (NetOffice.MSProjectApi.Enums.PjField)field;
			object newNewVal = (object)newVal;
			NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateObjectFromComProxy(_eventClass, info) as NetOffice.MSProjectApi.EventInfo;
			object[] paramsArray = new object[4];
			paramsArray[0] = newres;
			paramsArray[1] = newField;
			paramsArray[2] = newNewVal;
			paramsArray[3] = newInfo;
			_eventBinding.RaiseCustomEvent("ProjectBeforeResourceChange2", ref paramsArray);
		}

		public void ProjectBeforeAssignmentChange2([In, MarshalAs(UnmanagedType.IDispatch)] object asg, [In] object field, [In] object newVal, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectBeforeAssignmentChange2");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(asg, field, newVal, info);
				return;
			}

			NetOffice.MSProjectApi.Assignment newasg = Factory.CreateObjectFromComProxy(_eventClass, asg) as NetOffice.MSProjectApi.Assignment;
			NetOffice.MSProjectApi.Enums.PjAssignmentField newField = (NetOffice.MSProjectApi.Enums.PjAssignmentField)field;
			object newNewVal = (object)newVal;
			NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateObjectFromComProxy(_eventClass, info) as NetOffice.MSProjectApi.EventInfo;
			object[] paramsArray = new object[4];
			paramsArray[0] = newasg;
			paramsArray[1] = newField;
			paramsArray[2] = newNewVal;
			paramsArray[3] = newInfo;
			_eventBinding.RaiseCustomEvent("ProjectBeforeAssignmentChange2", ref paramsArray);
		}

		public void ProjectBeforeTaskNew2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectBeforeTaskNew2");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj, info);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateObjectFromComProxy(_eventClass, info) as NetOffice.MSProjectApi.EventInfo;
			object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray[1] = newInfo;
			_eventBinding.RaiseCustomEvent("ProjectBeforeTaskNew2", ref paramsArray);
		}

		public void ProjectBeforeResourceNew2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectBeforeResourceNew2");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj, info);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateObjectFromComProxy(_eventClass, info) as NetOffice.MSProjectApi.EventInfo;
			object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray[1] = newInfo;
			_eventBinding.RaiseCustomEvent("ProjectBeforeResourceNew2", ref paramsArray);
		}

		public void ProjectBeforeAssignmentNew2([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectBeforeAssignmentNew2");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj, info);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateObjectFromComProxy(_eventClass, info) as NetOffice.MSProjectApi.EventInfo;
			object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray[1] = newInfo;
			_eventBinding.RaiseCustomEvent("ProjectBeforeAssignmentNew2", ref paramsArray);
		}

		public void ApplicationBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object info)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ApplicationBeforeClose");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(info);
				return;
			}

			NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateObjectFromComProxy(_eventClass, info) as NetOffice.MSProjectApi.EventInfo;
			object[] paramsArray = new object[1];
			paramsArray[0] = newInfo;
			_eventBinding.RaiseCustomEvent("ApplicationBeforeClose", ref paramsArray);
		}

		public void OnUndoOrRedo([In] object bstrLabel, [In] object bstrGUID, [In] object fUndo)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("OnUndoOrRedo");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(bstrLabel, bstrGUID, fUndo);
				return;
			}

			string newbstrLabel = Convert.ToString(bstrLabel);
			string newbstrGUID = Convert.ToString(bstrGUID);
			bool newfUndo = Convert.ToBoolean(fUndo);
			object[] paramsArray = new object[3];
			paramsArray[0] = newbstrLabel;
			paramsArray[1] = newbstrGUID;
			paramsArray[2] = newfUndo;
			_eventBinding.RaiseCustomEvent("OnUndoOrRedo", ref paramsArray);
		}

		public void AfterCubeBuilt([In] [Out] ref object cubeFileName)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterCubeBuilt");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cubeFileName);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cubeFileName, 0);
			_eventBinding.RaiseCustomEvent("AfterCubeBuilt", ref paramsArray);

			cubeFileName = (string)paramsArray[0];
		}

		public void LoadWebPane([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In] [Out] ref object targetPage)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("LoadWebPane");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window, targetPage);
				return;
			}

			NetOffice.MSProjectApi.Window newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.MSProjectApi.Window;
			object[] paramsArray = new object[2];
			paramsArray[0] = newWindow;
			paramsArray.SetValue(targetPage, 1);
			_eventBinding.RaiseCustomEvent("LoadWebPane", ref paramsArray);

			targetPage = (string)paramsArray[1];
		}

		public void JobStart([In] object bstrName, [In] object bstrprojGuid, [In] object bstrjobGuid, [In] object jobType, [In] object lResult)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("JobStart");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(bstrName, bstrprojGuid, bstrjobGuid, jobType, lResult);
				return;
			}

			string newbstrName = Convert.ToString(bstrName);
			string newbstrprojGuid = Convert.ToString(bstrprojGuid);
			string newbstrjobGuid = Convert.ToString(bstrjobGuid);
			Int32 newjobType = Convert.ToInt32(jobType);
			Int32 newlResult = Convert.ToInt32(lResult);
			object[] paramsArray = new object[5];
			paramsArray[0] = newbstrName;
			paramsArray[1] = newbstrprojGuid;
			paramsArray[2] = newbstrjobGuid;
			paramsArray[3] = newjobType;
			paramsArray[4] = newlResult;
			_eventBinding.RaiseCustomEvent("JobStart", ref paramsArray);
		}

		public void JobCompleted([In] object bstrName, [In] object bstrprojGuid, [In] object bstrjobGuid, [In] object jobType, [In] object lResult)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("JobCompleted");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(bstrName, bstrprojGuid, bstrjobGuid, jobType, lResult);
				return;
			}

			string newbstrName = Convert.ToString(bstrName);
			string newbstrprojGuid = Convert.ToString(bstrprojGuid);
			string newbstrjobGuid = Convert.ToString(bstrjobGuid);
			Int32 newjobType = Convert.ToInt32(jobType);
			Int32 newlResult = Convert.ToInt32(lResult);
			object[] paramsArray = new object[5];
			paramsArray[0] = newbstrName;
			paramsArray[1] = newbstrprojGuid;
			paramsArray[2] = newbstrjobGuid;
			paramsArray[3] = newjobType;
			paramsArray[4] = newlResult;
			_eventBinding.RaiseCustomEvent("JobCompleted", ref paramsArray);
		}

		public void SaveStartingToServer([In] object bstrName, [In] object bstrprojGuid)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SaveStartingToServer");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(bstrName, bstrprojGuid);
				return;
			}

			string newbstrName = Convert.ToString(bstrName);
			string newbstrprojGuid = Convert.ToString(bstrprojGuid);
			object[] paramsArray = new object[2];
			paramsArray[0] = newbstrName;
			paramsArray[1] = newbstrprojGuid;
			_eventBinding.RaiseCustomEvent("SaveStartingToServer", ref paramsArray);
		}

		public void SaveCompletedToServer([In] object bstrName, [In] object bstrprojGuid)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SaveCompletedToServer");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(bstrName, bstrprojGuid);
				return;
			}

			string newbstrName = Convert.ToString(bstrName);
			string newbstrprojGuid = Convert.ToString(bstrprojGuid);
			object[] paramsArray = new object[2];
			paramsArray[0] = newbstrName;
			paramsArray[1] = newbstrprojGuid;
			_eventBinding.RaiseCustomEvent("SaveCompletedToServer", ref paramsArray);
		}

		public void ProjectBeforePublish([In, MarshalAs(UnmanagedType.IDispatch)] object pj, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProjectBeforePublish");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj, cancel);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			object[] paramsArray = new object[2];
			paramsArray[0] = newpj;
			paramsArray.SetValue(cancel, 1);
			_eventBinding.RaiseCustomEvent("ProjectBeforePublish", ref paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void PaneActivate()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PaneActivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("PaneActivate", ref paramsArray);
		}

		public void SecondaryViewChange([In, MarshalAs(UnmanagedType.IDispatch)] object window, [In, MarshalAs(UnmanagedType.IDispatch)] object prevView, [In, MarshalAs(UnmanagedType.IDispatch)] object newView, [In] object success)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SecondaryViewChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(window, prevView, newView, success);
				return;
			}

			NetOffice.MSProjectApi.Window newWindow = Factory.CreateObjectFromComProxy(_eventClass, window) as NetOffice.MSProjectApi.Window;
			NetOffice.MSProjectApi.View newprevView = Factory.CreateObjectFromComProxy(_eventClass, prevView) as NetOffice.MSProjectApi.View;
			NetOffice.MSProjectApi.View newnewView = Factory.CreateObjectFromComProxy(_eventClass, newView) as NetOffice.MSProjectApi.View;
			bool newsuccess = Convert.ToBoolean(success);
			object[] paramsArray = new object[4];
			paramsArray[0] = newWindow;
			paramsArray[1] = newprevView;
			paramsArray[2] = newnewView;
			paramsArray[3] = newsuccess;
			_eventBinding.RaiseCustomEvent("SecondaryViewChange", ref paramsArray);
		}

		public void IsFunctionalitySupported([In] object bstrFunctionality, [In, MarshalAs(UnmanagedType.IDispatch)] object info)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("IsFunctionalitySupported");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(bstrFunctionality, info);
				return;
			}

			string newbstrFunctionality = Convert.ToString(bstrFunctionality);
			NetOffice.MSProjectApi.EventInfo newInfo = Factory.CreateObjectFromComProxy(_eventClass, info) as NetOffice.MSProjectApi.EventInfo;
			object[] paramsArray = new object[2];
			paramsArray[0] = newbstrFunctionality;
			paramsArray[1] = newInfo;
			_eventBinding.RaiseCustomEvent("IsFunctionalitySupported", ref paramsArray);
		}

		public void ConnectionStatusChanged([In] object online)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ConnectionStatusChanged");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(online);
				return;
			}

			bool newonline = Convert.ToBoolean(online);
			object[] paramsArray = new object[1];
			paramsArray[0] = newonline;
			_eventBinding.RaiseCustomEvent("ConnectionStatusChanged", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}