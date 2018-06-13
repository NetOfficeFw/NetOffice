using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _ConversationHeader 
	/// SupportByVersion Outlook, 14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("00063107-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.OutlookApi.ConversationHeader))]
    public interface _ConversationHeader : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869453.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Application Application { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864386.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OutlookApi.Enums.OlObjectClass Class { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861297.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._NameSpace Session { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863021.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862787.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		string ConversationID { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864459.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		string ConversationTopic { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869015.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Conversation GetConversation();

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860331.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OutlookApi.SimpleItems GetItems();

		#endregion
	}
}
