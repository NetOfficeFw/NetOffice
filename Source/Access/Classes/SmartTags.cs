using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// CoClass SmartTags 
	/// SupportByVersion Access, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196170.aspx </remarks>
	[SupportByVersion("Access", 11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[TypeId("73778F0A-9743-4DF3-BBFA-941712488FEA")]
 	public interface SmartTags : _SmartTags
	{

	}
}
