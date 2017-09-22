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
    [ComImport, Guid("0006307C-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface OutlookBarShortcutsEvents
	{
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("newShortcut", typeof(OutlookApi.OutlookBarShortcut))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void ShortcutAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newShortcut);

		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61442)]
		void BeforeShortcutAdd([In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("shortcut", typeof(OutlookApi.OutlookBarShortcut))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61443)]
		void BeforeShortcutRemove([In, MarshalAs(UnmanagedType.IDispatch)] object shortcut, [In] [Out] ref object cancel);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class OutlookBarShortcutsEvents_SinkHelper : SinkHelper, OutlookBarShortcutsEvents
	{
		#region Static
		
		public static readonly string Id = "0006307C-0000-0000-C000-000000000046";
		
		#endregion
		
		#region Construction

		public OutlookBarShortcutsEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region OutlookBarShortcutsEvents
		
		public void ShortcutAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newShortcut)
		{
            if (!Validate("ShortcutAdd"))
            {
                Invoker.ReleaseParamsArray(newShortcut);
                return;
            }

			NetOffice.OutlookApi.OutlookBarShortcut newNewShortcut = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.OutlookBarShortcut>(EventClass, newShortcut, NetOffice.OutlookApi.OutlookBarShortcut.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newNewShortcut;
			EventBinding.RaiseCustomEvent("ShortcutAdd", ref paramsArray);
		}

		public void BeforeShortcutAdd([In] [Out] ref object cancel)
		{
            if (!Validate("BeforeShortcutAdd"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeforeShortcutAdd", ref paramsArray);

			cancel = ToBoolean(paramsArray[0]);
		}

		public void BeforeShortcutRemove([In, MarshalAs(UnmanagedType.IDispatch)] object shortcut, [In] [Out] ref object cancel)
        {
            if (!Validate("BeforeShortcutRemove"))
            {
                Invoker.ReleaseParamsArray(shortcut, cancel);
                return;
            }

			NetOffice.OutlookApi.OutlookBarShortcut newShortcut = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.OutlookBarShortcut>(EventClass, shortcut, NetOffice.OutlookApi.OutlookBarShortcut.LateBindingApiWrapperType);
			object[] paramsArray = new object[2];
			paramsArray[0] = newShortcut;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("BeforeShortcutRemove", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
        }

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}