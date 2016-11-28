using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;

namespace NetOffice.VBIDEApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
	[ComImport, Guid("0002E118-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _dispReferencesEvents
	{
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void ItemAdded([In, MarshalAs(UnmanagedType.IDispatch)] object reference);

		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void ItemRemoved([In, MarshalAs(UnmanagedType.IDispatch)] object reference);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _dispReferencesEvents_SinkHelper : SinkHelper, _dispReferencesEvents
	{
		#region Static
		
		public static readonly string Id = "0002E118-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public _dispReferencesEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

		#region _dispReferencesEvents Members
		
		public void ItemAdded([In, MarshalAs(UnmanagedType.IDispatch)] object reference)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ItemAdded");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(reference);
				return;
			}

			NetOffice.VBIDEApi.Reference newReference = Factory.CreateObjectFromComProxy(_eventClass, reference) as NetOffice.VBIDEApi.Reference;
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

			NetOffice.VBIDEApi.Reference newReference = Factory.CreateObjectFromComProxy(_eventClass, reference) as NetOffice.VBIDEApi.Reference;
			object[] paramsArray = new object[1];
			paramsArray[0] = newReference;
			_eventBinding.RaiseCustomEvent("ItemRemoved", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}