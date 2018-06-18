using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// DispatchInterface IVCharacters 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("000D0702-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.VisioApi.Characters))]
    public interface IVCharacters : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVApplication Application { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 Begin { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 CharCount { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="cellIndex">Int16 cellIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int16 get_CharProps(Int16 cellIndex);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="cellIndex">Int16 cellIndex</param>
        /// <param name="value">Int16 value</param>
        [SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		void set_CharProps(Int16 cellIndex, Int16 value);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CharProps
		/// </summary>
		/// <param name="cellIndex">Int16 cellIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_CharProps")]
		Int16 CharProps(Int16 cellIndex);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="biasLorR">Int16 biasLorR</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int16 get_CharPropsRow(Int16 biasLorR);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CharPropsRow
		/// </summary>
		/// <param name="biasLorR">Int16 biasLorR</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_CharPropsRow")]
		Int16 CharPropsRow(Int16 biasLorR);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 ObjectType { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 End { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 FieldCategory { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 FieldCode { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 FieldFormat { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string FieldFormula { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 IsField { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="cellIndex">Int16 cellIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int16 get_ParaProps(Int16 cellIndex);

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="cellIndex">Int16 cellIndex</param>
        /// <param name="value">Int16 value</param>
        [SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		void set_ParaProps(Int16 cellIndex, Int16 value);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ParaProps
		/// </summary>
		/// <param name="cellIndex">Int16 cellIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ParaProps")]
		Int16 ParaProps(Int16 cellIndex);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="biasLorR">Int16 biasLorR</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int16 get_ParaPropsRow(Int16 biasLorR);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ParaPropsRow
		/// </summary>
		/// <param name="biasLorR">Int16 biasLorR</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ParaPropsRow")]
		Int16 ParaPropsRow(Int16 biasLorR);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="biasLorR">Int16 biasLorR</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int16 get_TabPropsRow(Int16 biasLorR);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_TabPropsRow
		/// </summary>
		/// <param name="biasLorR">Int16 biasLorR</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_TabPropsRow")]
		Int16 TabPropsRow(Int16 biasLorR);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="runType">Int16 runType</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 get_RunBegin(Int16 runType);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_RunBegin
		/// </summary>
		/// <param name="runType">Int16 runType</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_RunBegin")]
		Int32 RunBegin(Int16 runType);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="runType">Int16 runType</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 get_RunEnd(Int16 runType);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_RunEnd
		/// </summary>
		/// <param name="runType">Int16 runType</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_RunEnd")]
		Int32 RunEnd(Int16 runType);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape Shape { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 Stat { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string TextAsString { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		object Text { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVDocument Document { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVEventList EventList { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 PersistsEvents { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string FieldFormulaU { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 ContainingPageID { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 ContainingMasterID { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="formula">string formula</param>
		/// <param name="format">Int16 format</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void AddCustomField(string formula, Int16 format);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="category">Int16 category</param>
		/// <param name="code">Int16 code</param>
		/// <param name="format">Int16 format</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void AddField(Int16 category, Int16 code, Int16 format);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Copy();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Cut();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Paste();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="formula">string formula</param>
		/// <param name="format">Int16 format</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void AddCustomFieldU(string formula, Int16 format);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Delete();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="category">NetOffice.VisioApi.Enums.VisFieldCategories category</param>
		/// <param name="code">NetOffice.VisioApi.Enums.VisFieldCodes code</param>
		/// <param name="format">NetOffice.VisioApi.Enums.VisFieldFormats format</param>
		/// <param name="langID">optional Int32 LangID = 0</param>
		/// <param name="calendarID">optional Int32 CalendarID = -1</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void AddFieldEx(NetOffice.VisioApi.Enums.VisFieldCategories category, NetOffice.VisioApi.Enums.VisFieldCodes code, NetOffice.VisioApi.Enums.VisFieldFormats format, object langID, object calendarID);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="category">NetOffice.VisioApi.Enums.VisFieldCategories category</param>
		/// <param name="code">NetOffice.VisioApi.Enums.VisFieldCodes code</param>
		/// <param name="format">NetOffice.VisioApi.Enums.VisFieldFormats format</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void AddFieldEx(NetOffice.VisioApi.Enums.VisFieldCategories category, NetOffice.VisioApi.Enums.VisFieldCodes code, NetOffice.VisioApi.Enums.VisFieldFormats format);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="category">NetOffice.VisioApi.Enums.VisFieldCategories category</param>
		/// <param name="code">NetOffice.VisioApi.Enums.VisFieldCodes code</param>
		/// <param name="format">NetOffice.VisioApi.Enums.VisFieldFormats format</param>
		/// <param name="langID">optional Int32 LangID = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void AddFieldEx(NetOffice.VisioApi.Enums.VisFieldCategories category, NetOffice.VisioApi.Enums.VisFieldCodes code, NetOffice.VisioApi.Enums.VisFieldFormats format, object langID);

		#endregion
	}
}
