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
    [ComImport, Guid("000D0B11-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface EDataRecordset
	{
		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("dataRecordsetChanged", typeof(VisioApi.IVDataRecordsetChangedEvent))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8224)]
		void DataRecordsetChanged([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordsetChanged);

		[SupportByVersion("Visio", 12,14,15,16)]
        [SinkArgument("dataRecordset", typeof(VisioApi.IVDataRecordset))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16416)]
		void BeforeDataRecordsetDelete([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordset);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class EDataRecordset_SinkHelper : SinkHelper, EDataRecordset
	{
		#region Static
		
		public static readonly string Id = "000D0B11-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

		public EDataRecordset_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region EDataRecordset
		
		public void DataRecordsetChanged([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordsetChanged)
        {
            if (!Validate("DataRecordsetChanged"))
            {
                Invoker.ReleaseParamsArray(dataRecordsetChanged);
                return;
            }

            NetOffice.VisioApi.IVDataRecordsetChangedEvent newDataRecordsetChanged = Factory.CreateEventArgumentObjectFromComProxy(EventClass, dataRecordsetChanged) as IVDataRecordsetChangedEvent;
            object[] paramsArray = new object[1];
			paramsArray[0] = newDataRecordsetChanged;
			EventBinding.RaiseCustomEvent("DataRecordsetChanged", ref paramsArray);
		}

		public void BeforeDataRecordsetDelete([In, MarshalAs(UnmanagedType.IDispatch)] object dataRecordset)
		{
            if (!Validate("BeforeDataRecordsetDelete"))
            {
                Invoker.ReleaseParamsArray(dataRecordset);
                return;
            }

            NetOffice.VisioApi.IVDataRecordset newDataRecordset = Factory.CreateEventArgumentObjectFromComProxy(EventClass, dataRecordset) as NetOffice.VisioApi.IVDataRecordset;
            object[] paramsArray = new object[1];
			paramsArray[0] = newDataRecordset;
			EventBinding.RaiseCustomEvent("BeforeDataRecordsetDelete", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}