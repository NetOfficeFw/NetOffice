using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// CoClass AllFunctions
	/// SupportByVersion Access, 10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845448.aspx </remarks>
	[SupportByVersion("Access", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[TypeId("08F6C822-3CFD-11D1-98BC-006008197D41")]
 	public interface AllFunctions : AllObjects
	{

	}
}
