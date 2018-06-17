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
	/// Default implementation of <see cref="NetOffice.AccessApi.EventContracts.DispEmptyCellEvents"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class DispEmptyCellEvents_SinkHelper : SinkHelper, NetOffice.AccessApi.EventContracts.DispEmptyCellEvents
	{
		#region Static
		
		/// <summary>
		/// Interface Id from DispEmptyCellEvents
		/// </summary>
		public static readonly string Id = "3B06E988-E47C-11CD-8701-00AA003F0F07";
		
		#endregion
	
		#region Ctor

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public DispEmptyCellEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion	

		#region DispEmptyCellEvents
		
		#endregion
	}
	
}
