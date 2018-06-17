using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// CoClass PageBreak 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844738.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._PageBreakEvents), typeof(EventContracts.DispPageBreakEvents))]
	[TypeId("3B06E95F-E47C-11CD-8701-00AA003F0F07")]
    public interface PageBreak : _PageBreak, IEventBinding
	{

	}
}
