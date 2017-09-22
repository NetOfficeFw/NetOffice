using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi.EventInterfaces
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("VBIDE", 12,14,5.3)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("0002E131-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _dispCommandBarControlEvents
	{
        [SinkArgument("commandBarControl", SinkArgumentType.UnknownProxy)]
        [SinkArgument("handled", SinkArgumentType.Bool)]
        [SinkArgument("cancelDefault", SinkArgumentType.Bool)]
        [SupportByVersion("VBIDE", 12,14,5.3)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void Click([In, MarshalAs(UnmanagedType.IDispatch)] object commandBarControl, [In] [Out] ref object handled, [In] [Out] ref object cancelDefault);
	}
	
	#endregion
	
	#region SinkHelper
	
    [InternalEntity(InternalEntityKind.SinkHelper)]
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _dispCommandBarControlEvents_SinkHelper : SinkHelper, _dispCommandBarControlEvents
	{
		#region Static
		
		public static readonly string Id = "0002E131-0000-0000-C000-000000000046";
		
		#endregion
		
		#region Ctor

		public _dispCommandBarControlEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region _dispCommandBarControlEvents
 
        public void Click([In, MarshalAs(UnmanagedType.IDispatch)] object commandBarControl, [In] [Out] ref object handled, [In] [Out] ref object cancelDefault)
		{
            if (!Validate("Click"))
            {
                Invoker.ReleaseParamsArray(commandBarControl, handled, cancelDefault);
                return;
            }

			object newCommandBarControl = Factory.CreateEventArgumentObjectFromComProxy(EventClass, commandBarControl) as object;
			object[] paramsArray = new object[3];
			paramsArray[0] = newCommandBarControl;
			paramsArray.SetValue(handled, 1);
			paramsArray.SetValue(cancelDefault, 2);
			EventBinding.RaiseCustomEvent("Click", ref paramsArray);

			handled = ToBoolean(paramsArray[1]);
			cancelDefault = ToBoolean(paramsArray[2]);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}