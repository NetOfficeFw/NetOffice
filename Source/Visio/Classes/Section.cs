using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Section_CellChangedEventHandler(NetOffice.VisioApi.IVCell cell);
	public delegate void Section_FormulaChangedEventHandler(NetOffice.VisioApi.IVCell cell);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Section 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff769382(v=office.14).aspx </remarks>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.VisioApi.EventContracts.ESection))]
	[TypeId("000D0A22-0000-0000-C000-000000000046")]
    public interface Section : IVSection, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765230(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Section_CellChangedEventHandler CellChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767144(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		event Section_FormulaChangedEventHandler FormulaChangedEvent;

		#endregion
	}
}
