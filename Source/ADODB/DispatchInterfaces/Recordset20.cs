using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface Recordset20 
	/// SupportByVersion ADODB, 2.1,2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.1,2.5)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("0000054F-0000-0010-8000-00AA006D2EA4")]
	public interface Recordset20 : Recordset15
	{
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5), ProxyResult]
		object DataSource { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5), ProxyResult]
		object ActiveCommand { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		bool StayInSync { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		string DataMember { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		void Cancel();

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// </summary>
		/// <param name="fileName">optional string fileName</param>
		/// <param name="persistFormat">optional NetOffice.ADODBApi.Enums.PersistFormatEnum PersistFormat = 0</param>
		[SupportByVersion("ADODB", 2.1)]
		void Save(object fileName, object persistFormat);

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1)]
		void Save();

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// </summary>
		/// <param name="fileName">optional string fileName</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1)]
		void Save(object fileName);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="stringFormat">optional NetOffice.ADODBApi.Enums.StringFormatEnum StringFormat = 2</param>
		/// <param name="numRows">optional Int32 NumRows = -1</param>
		/// <param name="columnDelimeter">optional string columnDelimeter</param>
		/// <param name="rowDelimeter">optional string rowDelimeter</param>
		/// <param name="nullExpr">optional string nullExpr</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		string GetString(object stringFormat, object numRows, object columnDelimeter, object rowDelimeter, object nullExpr);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		string GetString();

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="stringFormat">optional NetOffice.ADODBApi.Enums.StringFormatEnum StringFormat = 2</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		string GetString(object stringFormat);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="stringFormat">optional NetOffice.ADODBApi.Enums.StringFormatEnum StringFormat = 2</param>
		/// <param name="numRows">optional Int32 NumRows = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		string GetString(object stringFormat, object numRows);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="stringFormat">optional NetOffice.ADODBApi.Enums.StringFormatEnum StringFormat = 2</param>
		/// <param name="numRows">optional Int32 NumRows = -1</param>
		/// <param name="columnDelimeter">optional string columnDelimeter</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		string GetString(object stringFormat, object numRows, object columnDelimeter);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="stringFormat">optional NetOffice.ADODBApi.Enums.StringFormatEnum StringFormat = 2</param>
		/// <param name="numRows">optional Int32 NumRows = -1</param>
		/// <param name="columnDelimeter">optional string columnDelimeter</param>
		/// <param name="rowDelimeter">optional string rowDelimeter</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		string GetString(object stringFormat, object numRows, object columnDelimeter, object rowDelimeter);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="bookmark1">object bookmark1</param>
		/// <param name="bookmark2">object bookmark2</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		NetOffice.ADODBApi.Enums.CompareEnum CompareBookmarks(object bookmark1, object bookmark2);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="lockType">optional NetOffice.ADODBApi.Enums.LockTypeEnum LockType = -1</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		[BaseResult]
		NetOffice.ADODBApi._Recordset Clone(object lockType);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("ADODB", 2.1,2.5)]
        new NetOffice.ADODBApi._Recordset Clone();

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="affectRecords">optional NetOffice.ADODBApi.Enums.AffectEnum AffectRecords = 3</param>
		/// <param name="resyncValues">optional NetOffice.ADODBApi.Enums.ResyncEnum ResyncValues = 2</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		void Resync(object affectRecords, object resyncValues);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		void Resync();

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="affectRecords">optional NetOffice.ADODBApi.Enums.AffectEnum AffectRecords = 3</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		void Resync(object affectRecords);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="fileName">optional string fileName</param>
		/// <param name="persistFormat">optional NetOffice.ADODBApi.Enums.PersistFormatEnum PersistFormat = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("ADODB", 2.5)]
		void _xSave(object fileName, object persistFormat);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void _xSave();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="fileName">optional string fileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void _xSave(object fileName);

		#endregion
	}
}
