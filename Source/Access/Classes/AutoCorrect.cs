using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// CoClass AutoCorrect
	/// SupportByVersion Access, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822482.aspx </remarks>
	[SupportByVersion("Access", 11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[TypeId("4375351E-7052-40DF-B4D3-6095E7F8811B")]
 	public interface AutoCorrect : _AutoCorrect
	{

	}
}
