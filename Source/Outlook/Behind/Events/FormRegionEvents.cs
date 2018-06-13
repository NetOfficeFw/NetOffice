using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts.FormRegionEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class FormRegionEvents_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts.FormRegionEvents
	{
        #region Static

        /// <summary>
        /// Interface Id from FormRegionEvents
        /// </summary>
        public static readonly string Id = "0006305B-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
		public FormRegionEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region FormRegionEvents
		
        /// <summary>
        /// 
        /// </summary>
        /// <param name="expand"></param>
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

        /// <summary>
        /// 
        /// </summary>
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
}
