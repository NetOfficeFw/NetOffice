using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// DispatchInterface IVDataRecordset 
	/// SupportByVersion Visio, 12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("000D072F-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.VisioApi.DataRecordset))]
    public interface IVDataRecordset : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVApplication Application { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		Int16 Stat { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVDocument Document { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		Int16 ObjectType { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		Int32 ID { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		string Name { get; set; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		NetOffice.VisioApi.Enums.VisLinkReplaceBehavior LinkReplaceBehavior { get; set; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVDataConnection DataConnection { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVDataColumns DataColumns { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		string CommandString { get; set; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		string DataAsXML { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		DateTime TimeRefreshed { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		Int32 RefreshInterval { get; set; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		Int32 RefreshSettings { get; set; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVEventList EventList { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="primaryKeySettings">NetOffice.VisioApi.Enums.VisPrimaryKeySettings primaryKeySettings</param>
		/// <param name="primaryKey">String[] primaryKey</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		void GetPrimaryKey(out NetOffice.VisioApi.Enums.VisPrimaryKeySettings primaryKeySettings, out String[] primaryKey);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="primaryKeySettings">NetOffice.VisioApi.Enums.VisPrimaryKeySettings primaryKeySettings</param>
		/// <param name="primaryKey">String[] primaryKey</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		void SetPrimaryKey(NetOffice.VisioApi.Enums.VisPrimaryKeySettings primaryKeySettings, String[] primaryKey);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="criteriaString">string criteriaString</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		Int32[] GetDataRowIDs(string criteriaString);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="dataRowID">Int32 dataRowID</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		object[] GetRowData(Int32 dataRowID);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		void Delete();

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		void Refresh();

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="newDataAsXML">string newDataAsXML</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		void RefreshUsingXML(string newDataAsXML);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		NetOffice.VisioApi.IVShape[] GetAllRefreshConflicts();

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="shapeInConflict">NetOffice.VisioApi.IVShape shapeInConflict</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		void RemoveRefreshConflict(NetOffice.VisioApi.IVShape shapeInConflict);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="shapeInConflict">NetOffice.VisioApi.IVShape shapeInConflict</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		Int32[] GetMatchingRowsForRefreshConflict(NetOffice.VisioApi.IVShape shapeInConflict);

		#endregion
	}
}
