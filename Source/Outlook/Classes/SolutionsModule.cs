using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// CoClass SolutionsModule 
	/// SupportByVersion Outlook, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864298.aspx </remarks>
	[SupportByVersion("Outlook", 14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[TypeId("000610FF-0000-0000-C000-000000000046")]
 	public interface SolutionsModule : _SolutionsModule
	{

	}
}
