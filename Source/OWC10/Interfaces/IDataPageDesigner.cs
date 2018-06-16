using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// Interface IDataPageDesigner 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("F5B39ADD-1480-11D3-8549-00C04FAC67D7")]
	public interface IDataPageDesigner : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pDataSourceControl">NetOffice.OWC10Api.IDataSourceControl pDataSourceControl</param>
		[SupportByVersion("OWC10", 1)]
		Int32 ConnectDataComponents(NetOffice.OWC10Api.IDataSourceControl pDataSourceControl);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="sectType">NetOffice.OWC10Api.Enums.SectTypeEnum sectType</param>
		/// <param name="wzRecordsetName">string wzRecordsetName</param>
		[SupportByVersion("OWC10", 1)]
		Int32 CreateSection(NetOffice.OWC10Api.Enums.SectTypeEnum sectType, string wzRecordsetName);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="sectType">NetOffice.OWC10Api.Enums.SectTypeEnum sectType</param>
		/// <param name="wzRecordsetName">string wzRecordsetName</param>
		/// <param name="fInGroupingDefDelete">Int32 fInGroupingDefDelete</param>
		[SupportByVersion("OWC10", 1)]
		Int32 DeleteSection(NetOffice.OWC10Api.Enums.SectTypeEnum sectType, string wzRecordsetName, Int32 fInGroupingDefDelete);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pGroupLevel">NetOffice.OWC10Api.GroupLevel pGroupLevel</param>
		[SupportByVersion("OWC10", 1)]
		Int32 OnGroupLevelAdded(NetOffice.OWC10Api.GroupLevel pGroupLevel);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 OnGroupLevelDeleted();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pGroupLevel">NetOffice.OWC10Api.GroupLevel pGroupLevel</param>
		/// <param name="wzRecordsetNameOld">string wzRecordsetNameOld</param>
		/// <param name="wzRecordsetNameNew">string wzRecordsetNameNew</param>
		[SupportByVersion("OWC10", 1)]
		Int32 RebindGroupLevel(NetOffice.OWC10Api.GroupLevel pGroupLevel, string wzRecordsetNameOld, string wzRecordsetNameNew);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="wzConnectionString">string wzConnectionString</param>
		/// <param name="ppUnknownSharedConnection">object ppUnknownSharedConnection</param>
		[SupportByVersion("OWC10", 1)]
		Int32 GetSharedConnectionObject(string wzConnectionString, object ppUnknownSharedConnection);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="lMarker">Int32 lMarker</param>
		[SupportByVersion("OWC10", 1)]
		Int32 TWPerformanceMarker(Int32 lMarker);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="wzConnectionString">string wzConnectionString</param>
		/// <param name="pfSecure">Int32 pfSecure</param>
		[SupportByVersion("OWC10", 1)]
		Int32 IsDatabaseSecure(string wzConnectionString, Int32 pfSecure);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dispidChanged">Int32 dispidChanged</param>
		[SupportByVersion("OWC10", 1)]
		Int32 OnPropChanged(Int32 dispidChanged);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="wzConnectionString">string wzConnectionString</param>
		/// <param name="ppUnknownSharedDBNS">object ppUnknownSharedDBNS</param>
		[SupportByVersion("OWC10", 1)]
		Int32 GetSharedDBNS(string wzConnectionString, object ppUnknownSharedDBNS);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="ppbstrFileName">string ppbstrFileName</param>
		[SupportByVersion("OWC10", 1)]
		Int32 GetDatapagePath(string ppbstrFileName);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pfDesignMode">Int32 pfDesignMode</param>
		[SupportByVersion("OWC10", 1)]
		Int32 IsDesignMode(Int32 pfDesignMode);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pRequestingDSC">NetOffice.OWC10Api.IDataSourceControl pRequestingDSC</param>
		/// <param name="vfForceRefresh">bool vfForceRefresh</param>
		/// <param name="rt">NetOffice.OWC10Api.Enums.RefreshType rt</param>
		[SupportByVersion("OWC10", 1)]
		Int32 RefreshDataTools(NetOffice.OWC10Api.IDataSourceControl pRequestingDSC, bool vfForceRefresh, NetOffice.OWC10Api.Enums.RefreshType rt);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="ppbstrInstId">string ppbstrInstId</param>
		[SupportByVersion("OWC10", 1)]
		Int32 GetFieldListInstanceId(string ppbstrInstId);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pioum">NetOffice.OWC10Api.IOleUndoManager pioum</param>
		[SupportByVersion("OWC10", 1)]
		Int32 GetUndoManager(NetOffice.OWC10Api.IOleUndoManager pioum);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pDSC">NetOffice.OWC10Api.IDataSourceControl pDSC</param>
		/// <param name="bstrRecordSetDef">string bstrRecordSetDef</param>
		/// <param name="bstrDropRowsource">string bstrDropRowsource</param>
		/// <param name="varRowsources">object varRowsources</param>
		/// <param name="varRelationships">object varRelationships</param>
		/// <param name="ppprs">NetOffice.OWC10Api.PageRowsource ppprs</param>
		/// <param name="ppsr">NetOffice.OWC10Api.SchemaRelationship ppsr</param>
		[SupportByVersion("OWC10", 1)]
		Int32 DoRelWiz(NetOffice.OWC10Api.IDataSourceControl pDSC, string bstrRecordSetDef, string bstrDropRowsource, object varRowsources, object varRelationships, NetOffice.OWC10Api.PageRowsource ppprs, NetOffice.OWC10Api.SchemaRelationship ppsr);

		#endregion
	}
}
