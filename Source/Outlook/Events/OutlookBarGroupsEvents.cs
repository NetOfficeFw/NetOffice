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
    [ComImport, Guid("0006307B-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface OutlookBarGroupsEvents
	{
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("newGroup", typeof(OutlookApi.OutlookBarGroup))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void GroupAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newGroup);

		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61442)]
		void BeforeGroupAdd([In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("newGroup", typeof(OutlookApi.OutlookBarGroup))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61443)]
		void BeforeGroupRemove([In, MarshalAs(UnmanagedType.IDispatch)] object group, [In] [Out] ref object cancel);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class OutlookBarGroupsEvents_SinkHelper : SinkHelper, OutlookBarGroupsEvents
	{
		#region Static
		
		public static readonly string Id = "0006307B-0000-0000-C000-000000000046";
		
		#endregion

		#region Ctor

		public OutlookBarGroupsEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region OutlookBarGroupsEvents
		
		public void GroupAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newGroup)
		{
            if (!Validate("GroupAdd"))
            {
                Invoker.ReleaseParamsArray(newGroup);
                return;
            }

			NetOffice.OutlookApi.OutlookBarGroup newNewGroup = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.OutlookBarGroup>(EventClass, newGroup, NetOffice.OutlookApi.OutlookBarGroup.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newNewGroup;
			EventBinding.RaiseCustomEvent("GroupAdd", ref paramsArray);
		}

		public void BeforeGroupAdd([In] [Out] ref object cancel)
		{
            if (!Validate("BeforeGroupAdd"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeforeGroupAdd", ref paramsArray);

			cancel = ToBoolean(paramsArray[0]);
		}

		public void BeforeGroupRemove([In, MarshalAs(UnmanagedType.IDispatch)] object group, [In] [Out] ref object cancel)
        {
            if (!Validate("BeforeGroupRemove"))
            {
                Invoker.ReleaseParamsArray(group, cancel);
                return;
            }

			NetOffice.OutlookApi.OutlookBarGroup newGroup = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.OutlookBarGroup>(EventClass, group, NetOffice.OutlookApi.OutlookBarGroup.LateBindingApiWrapperType);
			object[] paramsArray = new object[2];
			paramsArray[0] = newGroup;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("BeforeGroupRemove", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}