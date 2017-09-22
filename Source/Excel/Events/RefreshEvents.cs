using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi.Events
{
#pragma warning disable

    #region SinkPoint Interface
   
    [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("0002441B-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface RefreshEvents
	{
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1596)]
		void BeforeRefresh([In] [Out] ref object cancel);

		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
        [SinkArgument("success", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1597)]
		void AfterRefresh([In] object success);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class RefreshEvents_SinkHelper : SinkHelper, RefreshEvents
	{
		#region Static
		
		public static readonly string Id = "0002441B-0000-0000-C000-000000000046";
		
		#endregion
			
		#region Ctor

		public RefreshEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region RefreshEvents
		
		public void BeforeRefresh([In] [Out] ref object cancel)
        {
            if (!Validate("BeforeRefresh"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeforeRefresh", ref paramsArray);

            cancel = ToBoolean(paramsArray[0]);
		}

		public void AfterRefresh([In] object success)
		{
            if (!Validate("AfterRefresh"))
            {
                Invoker.ReleaseParamsArray(success);
                return;
            }

			bool newSuccess = ToBoolean(success);
			object[] paramsArray = new object[1];
			paramsArray[0] = newSuccess;
			EventBinding.RaiseCustomEvent("AfterRefresh", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}