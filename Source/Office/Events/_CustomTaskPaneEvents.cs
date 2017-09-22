using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("Office", 12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000C033C-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _CustomTaskPaneEvents
	{
		[SupportByVersion("Office", 12,14,15,16)]
        [SinkArgument("customTaskPaneInst", typeof(NetOffice.OfficeApi._CustomTaskPane))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void VisibleStateChange([In, MarshalAs(UnmanagedType.IDispatch)] object customTaskPaneInst);

		[SupportByVersion("Office", 12,14,15,16)]
        [SinkArgument("customTaskPaneInst", typeof(NetOffice.OfficeApi._CustomTaskPane))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void DockPositionStateChange([In, MarshalAs(UnmanagedType.IDispatch)] object customTaskPaneInst);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _CustomTaskPaneEvents_SinkHelper : SinkHelper, _CustomTaskPaneEvents
	{
		#region Static
		
		public static readonly string Id = "000C033C-0000-0000-C000-000000000046";
		
		#endregion
		
		#region Ctor

		public _CustomTaskPaneEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region _CustomTaskPaneEvents
		
		public void VisibleStateChange([In, MarshalAs(UnmanagedType.IDispatch)] object customTaskPaneInst)
		{
            if (!Validate("VisibleStateChange"))
            {
                Invoker.ReleaseParamsArray(customTaskPaneInst);
                return;
            }

            NetOffice.OfficeApi._CustomTaskPane newCustomTaskPaneInst = Factory.CreateEventArgumentObjectFromComProxy(EventClass, customTaskPaneInst) as NetOffice.OfficeApi._CustomTaskPane;
			object[] paramsArray = new object[1];
			paramsArray[0] = newCustomTaskPaneInst;
			EventBinding.RaiseCustomEvent("VisibleStateChange", ref paramsArray);
        }

		public void DockPositionStateChange([In, MarshalAs(UnmanagedType.IDispatch)] object customTaskPaneInst)
		{
            if (!Validate("DockPositionStateChange"))
            {
                Invoker.ReleaseParamsArray(customTaskPaneInst);
                return;
            }

            NetOffice.OfficeApi._CustomTaskPane newCustomTaskPaneInst = Factory.CreateEventArgumentObjectFromComProxy(EventClass, customTaskPaneInst) as NetOffice.OfficeApi._CustomTaskPane;
            object[] paramsArray = new object[1];
			paramsArray[0] = newCustomTaskPaneInst;
            EventBinding.RaiseCustomEvent("DockPositionStateChange", ref paramsArray);
        }

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}