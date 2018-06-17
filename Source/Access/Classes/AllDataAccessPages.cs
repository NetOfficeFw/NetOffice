using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// CoClass AllDataAccessPages
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[TypeId("08F6C818-3CFD-11D1-98BC-006008197D41")]
 	public interface AllDataAccessPages : AllObjects
	{

	}
}
