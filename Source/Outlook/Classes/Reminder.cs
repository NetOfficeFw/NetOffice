using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// CoClass Reminder 
	/// SupportByVersion Outlook, 10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868711.aspx </remarks>
	[SupportByVersion("Outlook", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[TypeId("0006F028-0000-0000-C000-000000000046")]
 	public interface Reminder : _Reminder
	{

	}
}
