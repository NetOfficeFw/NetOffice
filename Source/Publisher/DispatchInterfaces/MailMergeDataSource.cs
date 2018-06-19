using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface MailMergeDataSource 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("1A0CD25D-4080-4A83-9DD9-02075743E446")]
	public interface MailMergeDataSource : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 ActiveRecord { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		string ConnectString { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.MailMergeDataFields DataFields { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.MailMergeFilters Filters { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 FirstRecord { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		bool Included { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		bool InvalidAddress { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		string InvalidComments { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 LastRecord { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.MailMergeMappedDataFields MappedDataFields { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		string Name { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 RecordCount { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 Type { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		string TableName { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.MailMergeDataSources DataSources { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		bool IsMaster { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		bool EverValidated { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		bool ValidatedClean { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="findText">string findText</param>
		/// <param name="field">optional string Field = </param>
		[SupportByVersion("Publisher", 14,15,16)]
		bool FindRecord(string findText, object field);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="findText">string findText</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		bool FindRecord(string findText);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="included">bool included</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void SetAllIncludedFlags(bool included);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="invalid">bool invalid</param>
		/// <param name="invalidComment">optional string InvalidComment = </param>
		[SupportByVersion("Publisher", 14,15,16)]
		void SetAllErrorFlags(bool invalid, object invalidComment);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="invalid">bool invalid</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void SetAllErrorFlags(bool invalid);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="sortField1">optional string SortField1 = </param>
		/// <param name="sortAscending1">optional bool SortAscending1 = true</param>
		/// <param name="sortField2">optional string SortField2 = </param>
		/// <param name="sortAscending2">optional bool SortAscending2 = true</param>
		/// <param name="sortField3">optional string SortField3 = </param>
		/// <param name="sortAscending3">optional bool SortAscending3 = true</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void SetSortOrder(object sortField1, object sortAscending1, object sortField2, object sortAscending2, object sortField3, object sortAscending3);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void SetSortOrder();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="sortField1">optional string SortField1 = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void SetSortOrder(object sortField1);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="sortField1">optional string SortField1 = </param>
		/// <param name="sortAscending1">optional bool SortAscending1 = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void SetSortOrder(object sortField1, object sortAscending1);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="sortField1">optional string SortField1 = </param>
		/// <param name="sortAscending1">optional bool SortAscending1 = true</param>
		/// <param name="sortField2">optional string SortField2 = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void SetSortOrder(object sortField1, object sortAscending1, object sortField2);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="sortField1">optional string SortField1 = </param>
		/// <param name="sortAscending1">optional bool SortAscending1 = true</param>
		/// <param name="sortField2">optional string SortField2 = </param>
		/// <param name="sortAscending2">optional bool SortAscending2 = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void SetSortOrder(object sortField1, object sortAscending1, object sortField2, object sortAscending2);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="sortField1">optional string SortField1 = </param>
		/// <param name="sortAscending1">optional bool SortAscending1 = true</param>
		/// <param name="sortField2">optional string SortField2 = </param>
		/// <param name="sortAscending2">optional bool SortAscending2 = true</param>
		/// <param name="sortField3">optional string SortField3 = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void SetSortOrder(object sortField1, object sortAscending1, object sortField2, object sortAscending2, object sortField3);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void OpenRecipientsDialog();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void ApplyFilter();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void Close();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="lRec">Int32 lRec</param>
		/// <param name="varField">object varField</param>
		/// <param name="value">object value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void EditRecord(Int32 lRec, object varField, object value);

		#endregion
	}
}
