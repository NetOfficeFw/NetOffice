using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// CoClass MacroError 
	/// SupportByVersion Access, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194165.aspx </remarks>
	[SupportByVersion("Access", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[TypeId("3F1B1773-65CB-4DB9-9FC6-ACED47DB285A")]
 	public interface MacroError : _MacroError
	{

	}
}
