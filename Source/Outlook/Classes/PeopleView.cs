using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// CoClass PeopleView 
	/// SupportByVersion Outlook, 15, 16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229577.aspx </remarks>
	[SupportByVersion("Outlook", 15, 16)]
	[EntityType(EntityType.IsCoClass)]
	[TypeId("0006200C-0000-0000-C000-000000000046")]
 	public interface PeopleView : _PeopleView
	{

	}
}
