using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts._DInspectorEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _DInspectorEvents_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts._DInspectorEvents
	{
        #region Static

        /// <summary>
        /// Interface Id from _DInspectorEvents
        /// </summary>
        public static readonly string Id = "2D9C6D57-BD3C-4275-BED2-73F0EDC18CCE";

        #endregion

        #region Ctor
        
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public _DInspectorEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region _DInspectorEvents Members
		
		#endregion
	}	
}
