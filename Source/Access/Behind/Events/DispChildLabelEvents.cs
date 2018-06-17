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
	/// Default implementation of <see cref="NetOffice.AccessApi.EventContracts.DispChildLabelEvents"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class DispChildLabelEvents_SinkHelper : SinkHelper, NetOffice.AccessApi.EventContracts.DispChildLabelEvents
	{
		#region Static
		
		/// <summary>
		/// Interface Id from DispChildLabelEvents
		/// </summary>
		public static readonly string Id = "2E705270-92D1-43CC-A57B-ED48BCCC711D";
		
		#endregion

		#region Ctor

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public DispChildLabelEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region DispChildLabelEvents
		
		#endregion
	}
	
}
