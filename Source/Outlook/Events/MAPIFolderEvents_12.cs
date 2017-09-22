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
    [ComImport, Guid("000630F7-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface MAPIFolderEvents_12
	{
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("moveTo", typeof(OutlookApi.MAPIFolder))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64424)]
		void BeforeFolderMove([In, MarshalAs(UnmanagedType.IDispatch)] object moveTo, [In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("item", SinkArgumentType.UnknownProxy)]
        [SinkArgument("moveTo", typeof(OutlookApi.MAPIFolder))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64425)]
		void BeforeItemMove([In, MarshalAs(UnmanagedType.IDispatch)] object item, [In, MarshalAs(UnmanagedType.IDispatch)] object moveTo, [In] [Out] ref object cancel);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class MAPIFolderEvents_12_SinkHelper : SinkHelper, MAPIFolderEvents_12
	{
		#region Static
		
		public static readonly string Id = "000630F7-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

		public MAPIFolderEvents_12_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region MAPIFolderEvents_12
		
		public void BeforeFolderMove([In, MarshalAs(UnmanagedType.IDispatch)] object moveTo, [In] [Out] ref object cancel)
		{
            if (!Validate("BeforeFolderMove"))
            {
                Invoker.ReleaseParamsArray(moveTo, cancel);
                return;
            }

            NetOffice.OutlookApi.MAPIFolder newMoveTo = Factory.CreateEventArgumentObjectFromComProxy(EventClass, moveTo) as NetOffice.OutlookApi.MAPIFolder;
            object[] paramsArray = new object[2];
			paramsArray[0] = newMoveTo;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("BeforeFolderMove", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
		}

		public void BeforeItemMove([In, MarshalAs(UnmanagedType.IDispatch)] object item, [In, MarshalAs(UnmanagedType.IDispatch)] object moveTo, [In] [Out] ref object cancel)
        {
            if (!Validate("BeforeItemMove"))
            {
                Invoker.ReleaseParamsArray(item, moveTo, cancel);
                return;
            }

			object newItem = Factory.CreateEventArgumentObjectFromComProxy(EventClass, item) as object;
            NetOffice.OutlookApi.MAPIFolder newMoveTo = Factory.CreateEventArgumentObjectFromComProxy(EventClass, moveTo) as NetOffice.OutlookApi.MAPIFolder;
            object[] paramsArray = new object[3];
			paramsArray[0] = newItem;
			paramsArray[1] = newMoveTo;
			paramsArray.SetValue(cancel, 2);
			EventBinding.RaiseCustomEvent("BeforeItemMove", ref paramsArray);

            cancel = ToBoolean(paramsArray[2]);
        }

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}