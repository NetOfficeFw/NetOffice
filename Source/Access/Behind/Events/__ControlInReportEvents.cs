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
	/// Default implementation of <see cref="NetOffice.AccessApi.EventContracts.__ControlInReportEvents"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class __ControlInReportEvents_SinkHelper : SinkHelper, NetOffice.AccessApi.EventContracts.__ControlInReportEvents
	{
		#region Static

		/// <summary>
		/// Interface Id from __ControlInReportEvents
		/// </summary>
		public static readonly string Id = "90B322A5-F1D9-11CD-8701-00AA003F0F07";

		#endregion

		#region Ctor

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public __ControlInReportEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}

		#endregion

		#region __ControlInReportEvents

		#endregion
	}
}
