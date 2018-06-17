using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// CoClass SmartTagActions 
	/// SupportByVersion Access, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194973.aspx </remarks>
	[SupportByVersion("Access", 11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[TypeId("AA533187-6399-4E6C-B6EC-6FC999E1C855")]
 	public interface SmartTagActions : _SmartTagActions
	{

	}
}
