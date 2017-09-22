using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("00063079-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface InspectorsEvents
	{
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void NewInspector([In, MarshalAs(UnmanagedType.IDispatch)] object inspector);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class InspectorsEvents_SinkHelper : SinkHelper, InspectorsEvents
	{
		#region Static
		
		public static readonly string Id = "00063079-0000-0000-C000-000000000046";
		
		#endregion

		#region Ctor

		public InspectorsEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region InspectorsEvents
		
		public void NewInspector([In, MarshalAs(UnmanagedType.IDispatch)] object inspector)
        {
            if (!Validate("NewInspector"))
            {
                Invoker.ReleaseParamsArray(inspector);
                return;
            }

            NetOffice.OutlookApi._Inspector newInspector = Factory.CreateEventArgumentObjectFromComProxy(EventClass, inspector) as NetOffice.OutlookApi._Inspector;
            object[] paramsArray = new object[1];
			paramsArray[0] = newInspector;
			EventBinding.RaiseCustomEvent("NewInspector", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}