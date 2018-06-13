using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _Conversation 
	/// SupportByVersion Outlook, 14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("00063101-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.OutlookApi.Conversation))]
    public interface _Conversation : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869259.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Application Application { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868054.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OutlookApi.Enums.OlObjectClass Class { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866390.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._NameSpace Session { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869565.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869792.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		string ConversationID { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866231.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OutlookApi.Table GetTable();

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868807.aspx </remarks>
		/// <param name="item">object item</param>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OutlookApi.SimpleItems GetChildren(object item);

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869780.aspx </remarks>
		/// <param name="item">object item</param>
		[SupportByVersion("Outlook", 14,15,16)]
		object GetParent(object item);

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866457.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OutlookApi.SimpleItems GetRootItems();

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869225.aspx </remarks>
		/// <param name="store">NetOffice.OutlookApi._Store store</param>
		[SupportByVersion("Outlook", 14,15,16)]
		string GetAlwaysAssignCategories(NetOffice.OutlookApi._Store store);

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867861.aspx </remarks>
		/// <param name="store">NetOffice.OutlookApi._Store store</param>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OutlookApi.Enums.OlAlwaysDeleteConversation GetAlwaysDelete(NetOffice.OutlookApi._Store store);

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869753.aspx </remarks>
		/// <param name="store">NetOffice.OutlookApi._Store store</param>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi.MAPIFolder GetAlwaysMoveToFolder(NetOffice.OutlookApi._Store store);

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867852.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		void MarkAsRead();

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868412.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		void MarkAsUnread();

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868084.aspx </remarks>
		/// <param name="categories">string categories</param>
		/// <param name="store">NetOffice.OutlookApi._Store store</param>
		[SupportByVersion("Outlook", 14,15,16)]
		void SetAlwaysAssignCategories(string categories, NetOffice.OutlookApi._Store store);

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869857.aspx </remarks>
		/// <param name="alwaysDelete">NetOffice.OutlookApi.Enums.OlAlwaysDeleteConversation alwaysDelete</param>
		/// <param name="store">NetOffice.OutlookApi._Store store</param>
		[SupportByVersion("Outlook", 14,15,16)]
		void SetAlwaysDelete(NetOffice.OutlookApi.Enums.OlAlwaysDeleteConversation alwaysDelete, NetOffice.OutlookApi._Store store);

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865038.aspx </remarks>
		/// <param name="moveToFolder">NetOffice.OutlookApi.MAPIFolder moveToFolder</param>
		/// <param name="store">NetOffice.OutlookApi._Store store</param>
		[SupportByVersion("Outlook", 14,15,16)]
		void SetAlwaysMoveToFolder(NetOffice.OutlookApi.MAPIFolder moveToFolder, NetOffice.OutlookApi._Store store);

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860425.aspx </remarks>
		/// <param name="store">NetOffice.OutlookApi._Store store</param>
		[SupportByVersion("Outlook", 14,15,16)]
		void ClearAlwaysAssignCategories(NetOffice.OutlookApi._Store store);

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869032.aspx </remarks>
		/// <param name="store">NetOffice.OutlookApi._Store store</param>
		[SupportByVersion("Outlook", 14,15,16)]
		void StopAlwaysDelete(NetOffice.OutlookApi._Store store);

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863707.aspx </remarks>
		/// <param name="store">NetOffice.OutlookApi._Store store</param>
		[SupportByVersion("Outlook", 14,15,16)]
		void StopAlwaysMoveToFolder(NetOffice.OutlookApi._Store store);

		#endregion
	}
}
