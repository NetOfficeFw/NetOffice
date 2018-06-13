using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _CardView 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("00063098-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.OutlookApi.CardView))]
    public interface _CardView : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867651.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Application Application { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869879.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlObjectClass Class { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862743.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._NameSpace Session { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861281.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865039.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		string Language { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868442.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		bool LockUserChanges { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869530.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		string Name { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866937.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlViewSaveOption SaveOption { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865870.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		bool Standard { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864001.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlViewType ViewType { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868292.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		string XML { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862760.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		string Filter { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865804.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.OrderFields SortFields { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869591.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		bool AllowInCellEditing { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869273.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		bool ShowEmptyFields { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865796.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		Int32 Width { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866427.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		Int32 MultiLineFieldHeight { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866268.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.ViewFields ViewFields { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863276.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.ViewFont HeadingsFont { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867479.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.ViewFont BodyFont { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862435.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.AutoFormatRules AutoFormatRules { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863041.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		void Apply();

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863578.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="saveOption">optional NetOffice.OutlookApi.Enums.OlViewSaveOption saveOption</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.View Copy(string name, object saveOption);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863578.aspx </remarks>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.View Copy(string name);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867650.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		void Delete();

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864421.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		void Reset();

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866781.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		void Save();

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869012.aspx </remarks>
		/// <param name="date">DateTime date</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		void GoToDate(DateTime date);

		#endregion
	}
}
