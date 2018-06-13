using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts._DPageWrapCtrlEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _DPageWrapCtrlEvents_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts._DPageWrapCtrlEvents
	{
        #region Static

        /// <summary>
        /// Interface Id from _DPageWrapCtrlEvents
        /// </summary>
        public static readonly string Id = "494F0971-DD96-11D2-AF70-006008AFF117";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public _DPageWrapCtrlEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region _DPageWrapCtrlEvents
		
		#endregion
	}
	
}
