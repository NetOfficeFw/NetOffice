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
    [ComImport, Guid("0006305B-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface FormRegionEvents
	{
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("expand", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64312)]
		void Expanded([In] object expand);

		[SupportByVersion("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61444)]
		void Close();
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class FormRegionEvents_SinkHelper : SinkHelper, FormRegionEvents
	{
		#region Static
		
		public static readonly string Id = "0006305B-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

		public FormRegionEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region FormRegionEvents
		
		public void Expanded([In] object expand)
        {
            if (!Validate("Expanded"))
            {
                Invoker.ReleaseParamsArray(expand);
                return;
            }

			bool newExpand = ToBoolean(expand);
			object[] paramsArray = new object[1];
			paramsArray[0] = newExpand;
			EventBinding.RaiseCustomEvent("Expanded", ref paramsArray);
		}

		public void Close()
		{
            if (!Validate("Close"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Close", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}