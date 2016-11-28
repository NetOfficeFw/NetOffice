using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;

namespace NetOffice.VisioApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersionAttribute("Visio", 12,14,15,16)]
	[ComImport, Guid("000D0B11-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface EDataRecordset
	{
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8224)]
		void DataRecordsetChanged([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordsetChanged);

		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16416)]
		void BeforeDataRecordsetDelete([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordset);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class EDataRecordset_SinkHelper : SinkHelper, EDataRecordset
	{
		#region Static
		
		public static readonly string Id = "000D0B11-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public EDataRecordset_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

		#region EDataRecordset Members
		
		public void DataRecordsetChanged([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordsetChanged)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DataRecordsetChanged");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(dataRecordsetChanged);
				return;
			}

			NetOffice.VisioApi.IVDataRecordsetChangedEvent newDataRecordsetChanged = Factory.CreateObjectFromComProxy(_eventClass, dataRecordsetChanged) as NetOffice.VisioApi.IVDataRecordsetChangedEvent;
			object[] paramsArray = new object[1];
			paramsArray[0] = newDataRecordsetChanged;
			_eventBinding.RaiseCustomEvent("DataRecordsetChanged", ref paramsArray);
		}

		public void BeforeDataRecordsetDelete([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordset)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeDataRecordsetDelete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(dataRecordset);
				return;
			}

			NetOffice.VisioApi.IVDataRecordset newDataRecordset = Factory.CreateObjectFromComProxy(_eventClass, dataRecordset) as NetOffice.VisioApi.IVDataRecordset;
			object[] paramsArray = new object[1];
			paramsArray[0] = newDataRecordset;
			_eventBinding.RaiseCustomEvent("BeforeDataRecordsetDelete", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}