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
    [ComImport, Guid("000630F3-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface NavigationPaneEvents_12
	{
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("currentModule", typeof(OutlookApi.NavigationModule))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64457)]
		void ModuleSwitch([In, MarshalAs(UnmanagedType.IDispatch)] object currentModule);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class NavigationPaneEvents_12_SinkHelper : SinkHelper, NavigationPaneEvents_12
	{
		#region Static
		
		public static readonly string Id = "000630F3-0000-0000-C000-000000000046";
		
		#endregion

		#region Ctor

		public NavigationPaneEvents_12_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region NavigationPaneEvents_12
		
		public void ModuleSwitch([In, MarshalAs(UnmanagedType.IDispatch)] object currentModule)
		{
            if (!Validate("SelectedChange"))
            {
                Invoker.ReleaseParamsArray(currentModule);
                return;
            }

			NetOffice.OutlookApi.NavigationModule newCurrentModule = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.NavigationModule>(EventClass, currentModule, NetOffice.OutlookApi.NavigationModule.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newCurrentModule;
			EventBinding.RaiseCustomEvent("ModuleSwitch", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}