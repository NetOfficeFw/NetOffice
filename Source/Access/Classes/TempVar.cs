using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// CoClass TempVar 
	/// SupportByVersion Access, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193475.aspx </remarks>
	[SupportByVersion("Access", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[TypeId("12DCE806-EA8A-46AA-88DF-C4486EDB78E3")]
 	public interface TempVar : _TempVar
	{

	}
}
