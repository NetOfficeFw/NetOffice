using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// CoClass TempVars 
	/// SupportByVersion Access, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821458.aspx </remarks>
	[SupportByVersion("Access", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[TypeId("6D40D9DE-2821-44A8-BAF3-8011E362CF59")]
 	public interface TempVars : _TempVars
	{

	}
}
