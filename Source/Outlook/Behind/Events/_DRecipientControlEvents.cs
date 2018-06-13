using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OutlookApi.EventContracts._DRecipientControlEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _DRecipientControlEvents_SinkHelper : SinkHelper, NetOffice.OutlookApi.EventContracts._DRecipientControlEvents
	{
        #region Static

        /// <summary>
        /// Interface Id from _DRecipientControlEvents
        /// </summary>
        public static readonly string Id = "D87E7E17-6897-11CE-A6C0-00AA00608FAA";
		
		#endregion
	
		#region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
		public _DRecipientControlEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region _DRecipientControlEvents
		
		#endregion
	}	
}
