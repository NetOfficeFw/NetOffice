using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts.NavigationPaneEvents_12"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class NavigationPaneEvents_12_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts.NavigationPaneEvents_12
	{
        #region Static

        /// <summary>
        /// Interface Id from NavigationPaneEvents_12
        /// </summary>
        public static readonly string Id = "000630F3-0000-0000-C000-000000000046";
		
		#endregion

		#region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
		public NavigationPaneEvents_12_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);		}
		
		#endregion

		#region NavigationPaneEvents_12
		
        /// <summary>
        /// 
        /// </summary>
        /// <param name="currentModule"></param>
		public void ModuleSwitch([In, MarshalAs(UnmanagedType.IDispatch)] object currentModule)
		{
            if (!Validate("SelectedChange"))
            {
                Invoker.ReleaseParamsArray(currentModule);
                return;
            }

			NetOffice.OutlookApi.NavigationModule newCurrentModule = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.NavigationModule>(EventClass, currentModule, typeof(NetOffice.OutlookApi.NavigationModule));
			object[] paramsArray = new object[1];
			paramsArray[0] = newCurrentModule;
			EventBinding.RaiseCustomEvent("ModuleSwitch", ref paramsArray);
		}

		#endregion
	}	
}

