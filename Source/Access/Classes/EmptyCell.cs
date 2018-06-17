using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// CoClass EmptyCell 
	/// SupportByVersion Access, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194884.aspx </remarks>
	[SupportByVersion("Access", 14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.DispEmptyCellEvents))]
	[TypeId("3B06E986-E47C-11CD-8701-00AA003F0F07")]
    public interface EmptyCell : _EmptyCell, IEventBinding
	{

	}
}
