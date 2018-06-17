using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// CoClass ReturnVar 
	/// SupportByVersion Access, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197108.aspx </remarks>
	[SupportByVersion("Access", 14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[TypeId("27B378D1-DAE2-48A5-BB40-A1C2BA02631D")]
 	public interface ReturnVar : _ReturnVar
	{

	}
}
