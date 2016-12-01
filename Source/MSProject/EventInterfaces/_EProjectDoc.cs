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
	[ComImport, Guid("F81DD3C0-5089-11CF-A49D-00AA00574C74"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _EProjectDoc
	{
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void Open([In, MarshalAs(UnmanagedType.IDispatch)] object pj);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void BeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pj);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void BeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object pj);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void BeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object pj);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)]
		void Calculate([In, MarshalAs(UnmanagedType.IDispatch)] object pj);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
		void Change([In, MarshalAs(UnmanagedType.IDispatch)] object pj);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)]
		void Activate([In, MarshalAs(UnmanagedType.IDispatch)] object pj);

		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)]
		void Deactivate([In, MarshalAs(UnmanagedType.IDispatch)] object pj);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _EProjectDoc_SinkHelper : SinkHelper, _EProjectDoc
	{
		#region Static
		
		public static readonly string Id = "F81DD3C0-5089-11CF-A49D-00AA00574C74";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public _EProjectDoc_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

		#region _EProjectDoc Members
		
		public void Open([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Open");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			_eventBinding.RaiseCustomEvent("Open", ref paramsArray);
		}

		public void BeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeClose");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			_eventBinding.RaiseCustomEvent("BeforeClose", ref paramsArray);
		}

		public void BeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeSave");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			_eventBinding.RaiseCustomEvent("BeforeSave", ref paramsArray);
		}

		public void BeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforePrint");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			_eventBinding.RaiseCustomEvent("BeforePrint", ref paramsArray);
		}

		public void Calculate([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Calculate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			_eventBinding.RaiseCustomEvent("Calculate", ref paramsArray);
		}

		public void Change([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Change");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			_eventBinding.RaiseCustomEvent("Change", ref paramsArray);
		}

		public void Activate([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Activate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			_eventBinding.RaiseCustomEvent("Activate", ref paramsArray);
		}

		public void Deactivate([In, MarshalAs(UnmanagedType.IDispatch)] object pj)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Deactivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pj);
				return;
			}

			NetOffice.MSProjectApi.Project newpj = Factory.CreateObjectFromComProxy(_eventClass, pj) as NetOffice.MSProjectApi.Project;
			object[] paramsArray = new object[1];
			paramsArray[0] = newpj;
			_eventBinding.RaiseCustomEvent("Deactivate", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}