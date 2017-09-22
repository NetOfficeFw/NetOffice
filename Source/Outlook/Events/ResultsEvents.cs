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

	[SupportByVersion("Outlook", 10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("0006300D-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ResultsEvents
	{
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("item", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void ItemAdd([In, MarshalAs(UnmanagedType.IDispatch)] object item);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("item", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61442)]
		void ItemChange([In, MarshalAs(UnmanagedType.IDispatch)] object item);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61443)]
		void ItemRemove();
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ResultsEvents_SinkHelper : SinkHelper, ResultsEvents
	{
		#region Static
		
		public static readonly string Id = "0006300D-0000-0000-C000-000000000046";
		
		#endregion
			
		#region Ctor

		public ResultsEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region ResultsEvents Members
		
		public void ItemAdd([In, MarshalAs(UnmanagedType.IDispatch)] object item)
        {
            if (!Validate("ItemAdd"))
            {
                Invoker.ReleaseParamsArray(item);
                return;
            }

			object newItem = Factory.CreateEventArgumentObjectFromComProxy(EventClass, item) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newItem;
			EventBinding.RaiseCustomEvent("ItemAdd", ref paramsArray);
		}

		public void ItemChange([In, MarshalAs(UnmanagedType.IDispatch)] object item)
		{
            if (!Validate("ItemChange"))
            {
                Invoker.ReleaseParamsArray(item);
                return;
            }

            object newItem = Factory.CreateEventArgumentObjectFromComProxy(EventClass, item) as object;
            object[] paramsArray = new object[1];
			paramsArray[0] = newItem;
			EventBinding.RaiseCustomEvent("ItemChange", ref paramsArray);
		}

		public void ItemRemove()
		{
            if (!Validate("ItemRemove"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("ItemRemove", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}