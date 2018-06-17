using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _AddTileUIReceiver 
	/// SupportByVersion Access, 15, 16
	/// </summary>
	[SupportByVersion("Access", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("1DD4E82D-9EF3-4730-A55E-4D179CB08006")]
	public interface _AddTileUIReceiver : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		string GetClientProtocolVersion();

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrTableName">string bstrTableName</param>
		/// <param name="bstrNounID">string bstrNounID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		string CreateCustomTable(string bstrTableName, string bstrNounID);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		string GetNounsVersion();

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		string GetNounsMetadata();

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrNounID">string bstrNounID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		string GetDefinitionOfNounID(string bstrNounID);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="pdispNounDefArray">object pdispNounDefArray</param>
		/// <param name="pdispFinalNameArray">object pdispFinalNameArray</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		void CreateObjectFromDefinition(object pdispNounDefArray, object pdispFinalNameArray);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrLeftTable">string bstrLeftTable</param>
		/// <param name="bstrRightTable">string bstrRightTable</param>
		/// <param name="bstrLookupFieldName">string bstrLookupFieldName</param>
		/// <param name="bstrLookupFieldDescription">string bstrLookupFieldDescription</param>
		/// <param name="lookupFieldPosition">Int32 lookupFieldPosition</param>
		/// <param name="iOptions">Int32 iOptions</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		void CreateRelationship(string bstrLeftTable, string bstrRightTable, string bstrLookupFieldName, string bstrLookupFieldDescription, Int32 lookupFieldPosition, Int32 iOptions);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="type">Int16 type</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		void ImportData(Int16 type);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		string GetNounTables();

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrSearchTerm">string bstrSearchTerm</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		void RegisterSearchTerm(string bstrSearchTerm);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		void BeginBatchNounAdd();

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		void FinishBatchNounAdd();

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="fVisible">bool fVisible</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		void NotifyAddTileUIVisibilityChange(bool fVisible);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="type">Int16 type</param>
		/// <param name="bstrUrl">string bstrUrl</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		void LaunchHyperlink(Int16 type, string bstrUrl);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		void MetadataLoaded();

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		bool IsOnlineContentAllowed();

		#endregion
	}
}
