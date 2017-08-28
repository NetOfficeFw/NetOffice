using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.VisioApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("Visio", 12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000D0B10-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface EDataRecordsets
	{
		[SupportByVersion("Visio", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(32800)]
		void DataRecordsetAdded([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordset);

		[SupportByVersion("Visio", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16416)]
		void BeforeDataRecordsetDelete([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordset);

		[SupportByVersion("Visio", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8224)]
		void DataRecordsetChanged([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordsetChanged);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class EDataRecordsets_SinkHelper : SinkHelper, EDataRecordsets
	{
		#region Static
		
		public static readonly string Id = "000D0B10-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private ICOMObject _eventClass;
        
		#endregion
		
		#region Construction

		public EDataRecordsets_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

		#region EDataRecordsets Members
		
		public void DataRecordsetAdded([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordset)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DataRecordsetAdded");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(dataRecordset);
				return;
			}

			NetOffice.VisioApi.IVDataRecordset newDataRecordset = Factory.CreateObjectFromComProxy(_eventClass, dataRecordset) as NetOffice.VisioApi.IVDataRecordset;
			object[] paramsArray = new object[1];
			paramsArray[0] = newDataRecordset;
			_eventBinding.RaiseCustomEvent("DataRecordsetAdded", ref paramsArray);
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

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}