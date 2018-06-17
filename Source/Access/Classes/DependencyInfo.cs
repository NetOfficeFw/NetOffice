using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// CoClass DependencyInfo 
	/// SupportByVersion Access, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193194.aspx </remarks>
	[SupportByVersion("Access", 11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[TypeId("F8C9DCB3-4063-490E-A73C-3533207CBC26")]
 	public interface DependencyInfo : _DependencyInfo
	{

	}
}
