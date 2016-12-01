using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;

namespace NetOffice.AccessApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
	[ComImport, Guid("F163F201-ADA2-11CF-89A9-00A0C9054129"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _References_Events
	{
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0)]
		void ItemAdded([In, MarshalAs(UnmanagedType.IDispatch)] object reference);

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void ItemRemoved([In, MarshalAs(UnmanagedType.IDispatch)] object reference);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _References_Events_SinkHelper : SinkHelper, _References_Events
	{
		#region Static
		
		public static readonly string Id = "F163F201-ADA2-11CF-89A9-00A0C9054129";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public _References_Events_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

		#region _References_Events Members
		
		public void ItemAdded([In, MarshalAs(UnmanagedType.IDispatch)] object reference)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ItemAdded");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(reference);
				return;
			}

			NetOffice.AccessApi.Reference newReference = Factory.CreateObjectFromComProxy(_eventClass, reference) as NetOffice.AccessApi.Reference;
			object[] paramsArray = new object[1];
			paramsArray[0] = newReference;
			_eventBinding.RaiseCustomEvent("ItemAdded", ref paramsArray);
		}

		public void ItemRemoved([In, MarshalAs(UnmanagedType.IDispatch)] object reference)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ItemRemoved");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(reference);
				return;
			}

			NetOffice.AccessApi.Reference newReference = Factory.CreateObjectFromComProxy(_eventClass, reference) as NetOffice.AccessApi.Reference;
			object[] paramsArray = new object[1];
			paramsArray[0] = newReference;
			_eventBinding.RaiseCustomEvent("ItemRemoved", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}