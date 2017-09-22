using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("Office", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000C0354-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _CommandBarComboBoxEvents
	{
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
        [SinkArgument("ctrl", typeof(OfficeApi.CommandBarComboBox))]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void Change([In, MarshalAs(UnmanagedType.IDispatch)] object ctrl);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _CommandBarComboBoxEvents_SinkHelper : SinkHelper, _CommandBarComboBoxEvents
	{
		#region Static
		
		public static readonly string Id = "000C0354-0000-0000-C000-000000000046";
		
		#endregion
			
		#region Ctor

		public _CommandBarComboBoxEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region _CommandBarComboBoxEvents
		
		public void Change([In, MarshalAs(UnmanagedType.IDispatch)] object ctrl)
        {
            if (!Validate("Change"))
            {
                Invoker.ReleaseParamsArray(ctrl);
                return;
            }

			NetOffice.OfficeApi.CommandBarComboBox newCtrl = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CommandBarComboBox>(EventClass, ctrl, NetOffice.OfficeApi.CommandBarComboBox.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newCtrl;
			EventBinding.RaiseCustomEvent("Change", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}