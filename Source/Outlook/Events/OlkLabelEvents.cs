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

	[SupportByVersion("Outlook", 12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000672E5-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface OlkLabelEvents
	{
		[SupportByVersion("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-600)]
		void Click();

		[SupportByVersion("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-601)]
		void DoubleClick();

		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Enum, typeof(OutlookApi.Enums.OlMouseButton))]
        [SinkArgument("shift", SinkArgumentType.Enum, typeof(OutlookApi.Enums.OlShiftState))]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-605)]
		void MouseDown([In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Enum, typeof(OutlookApi.Enums.OlMouseButton))]
        [SinkArgument("shift", SinkArgumentType.Enum, typeof(OutlookApi.Enums.OlShiftState))]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-606)]
		void MouseMove([In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("button", SinkArgumentType.Enum, typeof(OutlookApi.Enums.OlMouseButton))]
        [SinkArgument("shift", SinkArgumentType.Enum, typeof(OutlookApi.Enums.OlShiftState))]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-607)]
		void MouseUp([In] object button, [In] object shift, [In] object x, [In] object y);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class OlkLabelEvents_SinkHelper : SinkHelper, OlkLabelEvents
	{
		#region Static
		
		public static readonly string Id = "000672E5-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

		public OlkLabelEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region OlkLabelEvents
		
		public void Click()
		{
            if (!Validate("Click"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Click", ref paramsArray);
		}

		public void DoubleClick()
		{
            if (!Validate("DoubleClick"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("DoubleClick", ref paramsArray);
		}

		public void MouseDown([In] object button, [In] object shift, [In] object x, [In] object y)
		{
            if (!Validate("MouseDown"))
            {
                Invoker.ReleaseParamsArray(button, shift, x, y);
                return;
            }
           
			NetOffice.OutlookApi.Enums.OlMouseButton newButton = (NetOffice.OutlookApi.Enums.OlMouseButton)button;
			NetOffice.OutlookApi.Enums.OlShiftState newShift = (NetOffice.OutlookApi.Enums.OlShiftState)shift;
			Single newX = ToSingle(x);
			Single newY = ToSingle(y);
			object[] paramsArray = new object[4];
			paramsArray[0] = newButton;
			paramsArray[1] = newShift;
			paramsArray[2] = newX;
			paramsArray[3] = newY;
			EventBinding.RaiseCustomEvent("MouseDown", ref paramsArray);
		}

		public void MouseMove([In] object button, [In] object shift, [In] object x, [In] object y)
		{
            if (!Validate("MouseMove"))
            {
                Invoker.ReleaseParamsArray(button, shift, x, y);
                return;
            }
           
			NetOffice.OutlookApi.Enums.OlMouseButton newButton = (NetOffice.OutlookApi.Enums.OlMouseButton)button;
			NetOffice.OutlookApi.Enums.OlShiftState newShift = (NetOffice.OutlookApi.Enums.OlShiftState)shift;
			Single newX = ToSingle(x);
			Single newY = ToSingle(y);
			object[] paramsArray = new object[4];
			paramsArray[0] = newButton;
			paramsArray[1] = newShift;
			paramsArray[2] = newX;
			paramsArray[3] = newY;
			EventBinding.RaiseCustomEvent("MouseMove", ref paramsArray);
		}

		public void MouseUp([In] object button, [In] object shift, [In] object x, [In] object y)
		{
            if (!Validate("MouseUp"))
            {
                Invoker.ReleaseParamsArray(button, shift, x, y);
                return;
            }

            NetOffice.OutlookApi.Enums.OlMouseButton newButton = (NetOffice.OutlookApi.Enums.OlMouseButton)button;
			NetOffice.OutlookApi.Enums.OlShiftState newShift = (NetOffice.OutlookApi.Enums.OlShiftState)shift;
			Single newX = ToSingle(x);
			Single newY = ToSingle(y);
			object[] paramsArray = new object[4];
			paramsArray[0] = newButton;
			paramsArray[1] = newShift;
			paramsArray[2] = newX;
			paramsArray[3] = newY;
			EventBinding.RaiseCustomEvent("MouseUp", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}