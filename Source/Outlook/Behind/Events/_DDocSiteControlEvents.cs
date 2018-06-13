using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts._DDocSiteControlEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _DDocSiteControlEvents_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts._DDocSiteControlEvents
    {
        #region Static

        /// <summary>
        /// Interface Id from _DDocSiteControlEvents
        /// </summary>
        public static readonly string Id = "50BB9B50-811D-11CE-B565-00AA00608FAA";
		
		#endregion
		
		#region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
		public _DDocSiteControlEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region _DDocSiteControlEvents Members
		
		#endregion
	}	
}
