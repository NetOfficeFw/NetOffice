using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _PeopleView 
	/// SupportByVersion Outlook, 15, 16
	/// </summary>
	[SupportByVersion("Outlook", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("000630A3-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.OutlookApi.PeopleView))]
    public interface _PeopleView : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228406.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		[BaseResult]
		NetOffice.OutlookApi._Application Application { get; }

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230583.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		NetOffice.OutlookApi.Enums.OlObjectClass Class { get; }

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228591.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		[BaseResult]
		NetOffice.OutlookApi._NameSpace Session { get; }

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230405.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227687.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		string Language { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228057.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		bool LockUserChanges { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231136.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		string Name { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230036.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		NetOffice.OutlookApi.Enums.OlViewSaveOption SaveOption { get; }

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229062.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		bool Standard { get; }

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229669.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		NetOffice.OutlookApi.Enums.OlViewType ViewType { get; }

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228325.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		string XML { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228064.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		string Filter { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229733.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		NetOffice.OutlookApi.OrderFields SortFields { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227495.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		void Apply();

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231247.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="saveOption">optional NetOffice.OutlookApi.Enums.OlViewSaveOption saveOption</param>
		[SupportByVersion("Outlook", 15, 16)]
		NetOffice.OutlookApi.View Copy(string name, object saveOption);

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231247.aspx </remarks>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 15, 16)]
		NetOffice.OutlookApi.View Copy(string name);

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227780.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		void Delete();

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231541.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		void Reset();

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230523.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		void Save();

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230336.aspx </remarks>
		/// <param name="date">DateTime date</param>
		[SupportByVersion("Outlook", 15, 16)]
		void GoToDate(DateTime date);

		#endregion
	}
}
