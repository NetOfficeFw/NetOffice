using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _ContactItem 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("00063021-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.OutlookApi.ContactItem))]
    public interface _ContactItem : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868476.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Application Application { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866948.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlObjectClass Class { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868700.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._NameSpace Session { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864478.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862161.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.Actions Actions { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865274.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.Attachments Attachments { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868961.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string BillingInformation { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865615.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Body { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868930.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Categories { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863616.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Companies { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866064.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string ConversationIndex { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868608.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string ConversationTopic { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867864.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		DateTime CreationTime { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860429.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string EntryID { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865802.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.FormDescription FormDescription { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869254.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Inspector GetInspector { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869965.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlImportance Importance { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861547.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		DateTime LastModificationTime { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object MAPIOBJECT { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863901.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string MessageClass { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867548.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Mileage { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869251.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		bool NoAging { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864258.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		Int32 OutlookInternalVersion { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862406.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string OutlookVersion { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869811.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		bool Saved { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869844.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlSensitivity Sensitivity { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863365.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		Int32 Size { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868050.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Subject { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865670.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		bool UnRead { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869955.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.UserProperties UserProperties { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860971.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Account { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868919.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		DateTime Anniversary { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860640.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string AssistantName { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860982.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string AssistantTelephoneNumber { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869282.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		DateTime Birthday { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868769.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Business2TelephoneNumber { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867232.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string BusinessAddress { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866226.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string BusinessAddressCity { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869143.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string BusinessAddressCountry { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860952.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string BusinessAddressPostalCode { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864267.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string BusinessAddressPostOfficeBox { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860975.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string BusinessAddressState { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862080.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string BusinessAddressStreet { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867303.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string BusinessFaxNumber { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867885.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string BusinessHomePage { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866227.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string BusinessTelephoneNumber { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862663.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string CallbackTelephoneNumber { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864290.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string CarTelephoneNumber { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869537.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Children { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868047.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string CompanyAndFullName { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869476.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string CompanyLastFirstNoSpace { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867631.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string CompanyLastFirstSpaceOnly { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862248.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string CompanyMainTelephoneNumber { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860654.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string CompanyName { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863063.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string ComputerNetworkName { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867329.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string CustomerID { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866006.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Department { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860936.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Email1Address { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869942.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Email1AddressType { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866426.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Email1DisplayName { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867187.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Email1EntryID { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861590.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Email2Address { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860739.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Email2AddressType { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863592.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Email2DisplayName { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860946.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Email2EntryID { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868579.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Email3Address { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868560.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Email3AddressType { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869715.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Email3DisplayName { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869918.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Email3EntryID { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866001.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string FileAs { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863968.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string FirstName { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862471.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string FTPSite { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863059.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string FullName { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867814.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string FullNameAndCompany { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860335.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlGender Gender { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869210.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string GovernmentIDNumber { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865840.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Hobby { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861802.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Home2TelephoneNumber { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869043.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string HomeAddress { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862079.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string HomeAddressCity { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868311.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string HomeAddressCountry { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862692.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string HomeAddressPostalCode { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868107.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string HomeAddressPostOfficeBox { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868798.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string HomeAddressState { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868070.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string HomeAddressStreet { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869795.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string HomeFaxNumber { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869394.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string HomeTelephoneNumber { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869869.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Initials { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868657.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string InternetFreeBusyAddress { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868030.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string ISDNNumber { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866080.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string JobTitle { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863618.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		bool Journal { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866054.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Language { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868609.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string LastFirstAndSuffix { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862992.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string LastFirstNoSpace { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865046.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string LastFirstNoSpaceCompany { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868471.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string LastFirstSpaceOnly { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867834.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string LastFirstSpaceOnlyCompany { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864221.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string LastName { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866735.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string LastNameAndFirstName { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866926.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string MailingAddress { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870077.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string MailingAddressCity { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860947.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string MailingAddressCountry { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868831.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string MailingAddressPostalCode { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868668.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string MailingAddressPostOfficeBox { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868200.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string MailingAddressState { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867259.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string MailingAddressStreet { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868867.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string ManagerName { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860667.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string MiddleName { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864197.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string MobileTelephoneNumber { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869794.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string NetMeetingAlias { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867383.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string NetMeetingServer { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869407.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string NickName { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870105.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string OfficeLocation { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863903.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string OrganizationalIDNumber { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861601.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string OtherAddress { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868472.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string OtherAddressCity { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869088.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string OtherAddressCountry { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868437.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string OtherAddressPostalCode { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867645.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string OtherAddressPostOfficeBox { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868392.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string OtherAddressState { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869474.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string OtherAddressStreet { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868199.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string OtherFaxNumber { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862246.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string OtherTelephoneNumber { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862776.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string PagerNumber { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869119.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string PersonalHomePage { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868841.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string PrimaryTelephoneNumber { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864488.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Profession { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861307.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string RadioTelephoneNumber { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860437.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string ReferredBy { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867114.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlMailingAddress SelectedMailingAddress { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864728.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Spouse { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869774.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Suffix { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869871.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string TelexNumber { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863910.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Title { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867390.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string TTYTDDTelephoneNumber { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869725.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string User1 { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865797.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string User2 { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870180.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string User3 { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867668.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string User4 { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string UserCertificate { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860704.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string WebPage { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862400.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string YomiCompanyName { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868455.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string YomiFirstName { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864219.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string YomiLastName { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.Links Links { get; }

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864274.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		NetOffice.OutlookApi.ItemProperties ItemProperties { get; }

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861577.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		string LastFirstNoSpaceAndSuffix { get; }

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863071.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlDownloadState DownloadState { get; }

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869375.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		string IMAddress { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867597.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlRemoteStatus MarkForDownload { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863421.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		bool IsConflict { get; }

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869858.aspx </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		bool AutoResolvedWinner { get; }

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868985.aspx </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		NetOffice.OutlookApi.Conflicts Conflicts { get; }

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865639.aspx </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		bool HasPicture { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869642.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.PropertyAccessor PropertyAccessor { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870041.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		string TaskSubject { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863344.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		DateTime TaskDueDate { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870049.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		DateTime TaskStartDate { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865987.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		DateTime TaskCompletedDate { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860673.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		DateTime ToDoTaskOrdinal { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860694.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		bool ReminderOverrideDefault { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868429.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		bool ReminderPlaySound { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867607.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		bool ReminderSet { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868467.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		string ReminderSoundFile { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869058.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		DateTime ReminderTime { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868859.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		bool IsMarkedAsTask { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860753.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		string BusinessCardLayoutXml { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865308.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlBusinessCardType BusinessCardType { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861318.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		string ConversationID { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870062.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		object RTFBody { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861625.aspx </remarks>
		/// <param name="saveMode">NetOffice.OutlookApi.Enums.OlInspectorClose saveMode</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void Close(NetOffice.OutlookApi.Enums.OlInspectorClose saveMode);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861014.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		object Copy();

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862373.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void Delete();

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866773.aspx </remarks>
		/// <param name="modal">optional object modal</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void Display(object modal);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866773.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void Display();

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869633.aspx </remarks>
		/// <param name="destFldr">NetOffice.OutlookApi.MAPIFolder destFldr</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		object Move(NetOffice.OutlookApi.MAPIFolder destFldr);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867897.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void PrintOut();

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862148.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void Save();

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868220.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="type">optional object type</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void SaveAs(string path, object type);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868220.aspx </remarks>
		/// <param name="path">string path</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void SaveAs(string path);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863898.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.MailItem ForwardAsVcard();

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862359.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		void ShowCategoriesDialog();

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868443.aspx </remarks>
		/// <param name="path">string path</param>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		void AddPicture(string path);

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868359.aspx </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		void RemovePicture();

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863014.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.MailItem ForwardAsBusinessCard();

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867880.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		void ShowBusinessCardEditor();

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867386.aspx </remarks>
		/// <param name="path">string path</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		void SaveBusinessCardImage(string path);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863935.aspx </remarks>
		/// <param name="phoneNumber">NetOffice.OutlookApi.Enums.OlContactPhoneNumber phoneNumber</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		void ShowCheckPhoneDialog(NetOffice.OutlookApi.Enums.OlContactPhoneNumber phoneNumber);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869513.aspx </remarks>
		/// <param name="markInterval">NetOffice.OutlookApi.Enums.OlMarkInterval markInterval</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		void MarkAsTask(NetOffice.OutlookApi.Enums.OlMarkInterval markInterval);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861851.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		void ClearTaskFlag();

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868370.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		void ResetBusinessCard();

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866480.aspx </remarks>
		/// <param name="path">string path</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		void AddBusinessCardLogoPicture(string path);

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861831.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Conversation GetConversation();

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231065.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		void ShowCheckFullNameDialog();

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229480.aspx </remarks>
		/// <param name="mailingAddress">NetOffice.OutlookApi.Enums.OlMailingAddress mailingAddress</param>
		[SupportByVersion("Outlook", 15, 16)]
		void ShowCheckAddressDialog(NetOffice.OutlookApi.Enums.OlMailingAddress mailingAddress);

		#endregion
	}
}
