using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;
using NetOffice.Exceptions;

namespace NetOffice.AccessApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.AccessApi.EventContracts._DispControlInReportEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _DispControlInReportEvents_SinkHelper : SinkHelper, NetOffice.AccessApi.EventContracts._DispControlInReportEvents
	{
		#region Static
		
		/// <summary>
		/// Interface Id from _DispControlInReportEvents
		/// </summary>
		public static readonly string Id = "2E70527D-92D1-43CC-A57B-ED48BCCC711D";
		
		#endregion
			
		#region Ctor

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public _DispControlInReportEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region _DispControlInReportEvents Members
		
		#endregion
	}
	
}
