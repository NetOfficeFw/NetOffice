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
    [ComImport, Guid("0006307A-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface OutlookBarPaneEvents
	{
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("shortcut", typeof(OutlookApi.OutlookBarShortcut))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void BeforeNavigate([In, MarshalAs(UnmanagedType.IDispatch)] object shortcut, [In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("toGroup", typeof(OutlookApi.OutlookBarGroup))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61442)]
		void BeforeGroupSwitch([In, MarshalAs(UnmanagedType.IDispatch)] object toGroup, [In] [Out] ref object cancel);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class OutlookBarPaneEvents_SinkHelper : SinkHelper, OutlookBarPaneEvents
	{
		#region Static
		
		public static readonly string Id = "0006307A-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Construction

		public OutlookBarPaneEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region OutlookBarPaneEvents
		
		public void BeforeNavigate([In, MarshalAs(UnmanagedType.IDispatch)] object shortcut, [In] [Out] ref object cancel)
        {
            if (!Validate("BeforeNavigate"))
            {
                Invoker.ReleaseParamsArray(shortcut, cancel);
                return;
            }

			NetOffice.OutlookApi.OutlookBarShortcut newShortcut = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.OutlookBarShortcut>(EventClass, shortcut, NetOffice.OutlookApi.OutlookBarShortcut.LateBindingApiWrapperType);
			object[] paramsArray = new object[2];
			paramsArray[0] = newShortcut;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("BeforeNavigate", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
		}

		public void BeforeGroupSwitch([In, MarshalAs(UnmanagedType.IDispatch)] object toGroup, [In] [Out] ref object cancel)
        {
            if (!Validate("BeforeGroupSwitch"))
            {
                Invoker.ReleaseParamsArray(toGroup, cancel);
                return;
            }

			NetOffice.OutlookApi.OutlookBarGroup newToGroup = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.OutlookBarGroup>(EventClass, toGroup, NetOffice.OutlookApi.OutlookBarGroup.LateBindingApiWrapperType);
			object[] paramsArray = new object[2];
			paramsArray[0] = newToGroup;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("BeforeGroupSwitch", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
        }

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}