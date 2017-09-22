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
    [ComImport, Guid("00063078-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ExplorersEvents
	{
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("explorer", typeof(NetOffice.OutlookApi._Explorer))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void NewExplorer([In, MarshalAs(UnmanagedType.IDispatch)] object explorer);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ExplorersEvents_SinkHelper : SinkHelper, ExplorersEvents
	{
		#region Static
		
		public static readonly string Id = "00063078-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

		public ExplorersEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region ExplorersEvents
		
		public void NewExplorer([In, MarshalAs(UnmanagedType.IDispatch)] object explorer)
		{
            if (!Validate("NewExplorer"))
            {
                Invoker.ReleaseParamsArray(explorer);
                return;
            }

			NetOffice.OutlookApi._Explorer newExplorer = Factory.CreateEventArgumentObjectFromComProxy(EventClass, explorer) as NetOffice.OutlookApi._Explorer;
			object[] paramsArray = new object[1];
			paramsArray[0] = newExplorer;
			EventBinding.RaiseCustomEvent("NewExplorer", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}