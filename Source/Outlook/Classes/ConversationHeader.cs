using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// CoClass ConversationHeader 
	/// SupportByVersion Outlook, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865006.aspx </remarks>
	[SupportByVersion("Outlook", 14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[TypeId("00061107-0000-0000-C000-000000000046")]
 	public interface ConversationHeader : _ConversationHeader
	{

	}
}
