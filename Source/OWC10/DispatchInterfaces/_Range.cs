using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OWC10Api
{
    /// <summary>
    /// _Range
    /// </summary>
    [SyntaxBypass]
    public interface _Range_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.OWC10Api.Enums.XlReferenceStyle referenceStyle</param>
        /// <param name="external">optional object external</param>
        /// <param name="relativeTo">optional object relativeTo</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external, object relativeTo);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_Address
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.OWC10Api.Enums.XlReferenceStyle referenceStyle</param>
        /// <param name="external">optional object external</param>
        /// <param name="relativeTo">optional object relativeTo</param>
        [SupportByVersion("OWC10", 1), Redirect("get_Address")]
        string Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external, object relativeTo);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_Address(object rowAbsolute);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_Address
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        [SupportByVersion("OWC10", 1), Redirect("get_Address")]
        string Address(object rowAbsolute);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_Address(object rowAbsolute, object columnAbsolute);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_Address
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        [SupportByVersion("OWC10", 1), Redirect("get_Address")]
        string Address(object rowAbsolute, object columnAbsolute);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.OWC10Api.Enums.XlReferenceStyle referenceStyle</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_Address(object rowAbsolute, object columnAbsolute, object referenceStyle);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_Address
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.OWC10Api.Enums.XlReferenceStyle referenceStyle</param>
        [SupportByVersion("OWC10", 1), Redirect("get_Address")]
        string Address(object rowAbsolute, object columnAbsolute, object referenceStyle);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.OWC10Api.Enums.XlReferenceStyle referenceStyle</param>
        /// <param name="external">optional object external</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_Address
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.OWC10Api.Enums.XlReferenceStyle referenceStyle</param>
        /// <param name="external">optional object external</param>
        [SupportByVersion("OWC10", 1), Redirect("get_Address")]
        string Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="rowOffset">optional object rowOffset</param>
        /// <param name="columnOffset">optional object columnOffset</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OWC10Api._Range get_Offset(object rowOffset, object columnOffset);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_Offset
        /// </summary>
        /// <param name="rowOffset">optional object rowOffset</param>
        /// <param name="columnOffset">optional object columnOffset</param>
        [SupportByVersion("OWC10", 1), Redirect("get_Offset")]
        NetOffice.OWC10Api._Range Offset(object rowOffset, object columnOffset);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="rowOffset">optional object rowOffset</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OWC10Api._Range get_Offset(object rowOffset);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_Offset
        /// </summary>
        /// <param name="rowOffset">optional object rowOffset</param>
        [SupportByVersion("OWC10", 1), Redirect("get_Offset")]
        NetOffice.OWC10Api._Range Offset(object rowOffset);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        /// <param name="rangeValueDataType">optional object rangeValueDataType</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_Value(object rangeValueDataType);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        /// <param name="rangeValueDataType">optional object rangeValueDataType</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_Value(object rangeValueDataType, object value);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_Value
        /// </summary>
        /// <param name="rangeValueDataType">optional object rangeValueDataType</param>
        [SupportByVersion("OWC10", 1), Redirect("get_Value")]
        object Value(object rangeValueDataType);

        #endregion
    }

    /// <summary>
    /// DispatchInterface _Range 
    /// SupportByVersion OWC10, 1
    /// </summary>
    [SupportByVersion("OWC10", 1)]
    [EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "OWC10", 1), HasIndexProperty(IndexInvoke.Property, "_Default")]
	[TypeId("F5B39B05-1480-11D3-8549-00C04FAC67D7")]
    [CoClassSource(typeof(NetOffice.OWC10Api.Range))]
    public interface _Range : _Range_, IEnumerableProvider<object>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// Custom Indexer
        /// </summary>
        /// <param name="row">optional object row</param>
        [SupportByVersion("OWC10", 1)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty, CustomIndexer]
        object this[object row] { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        /// <param name="row">optional object row</param>
        /// <param name="column">optional object column</param>
        [SupportByVersion("OWC10", 1)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        object this[object row, object column] { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        new string Address { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        NetOffice.OWC10Api.ISpreadsheet Application { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.Borders Borders { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        NetOffice.OWC10Api._Range Cells { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        Int32 Column { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        NetOffice.OWC10Api._Range Columns { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        object ColumnWidth { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        Int32 Count { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        NetOffice.OWC10Api._Range CurrentArray { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        NetOffice.OWC10Api._Range CurrentRegion { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="direction">NetOffice.OWC10Api.Enums.XlDirection direction</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OWC10Api._Range get_End(NetOffice.OWC10Api.Enums.XlDirection direction);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_End
        /// </summary>
        /// <param name="direction">NetOffice.OWC10Api.Enums.XlDirection direction</param>
        [SupportByVersion("OWC10", 1), Redirect("get_End")]
        NetOffice.OWC10Api._Range End(NetOffice.OWC10Api.Enums.XlDirection direction);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        NetOffice.OWC10Api._Range EntireColumn { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        NetOffice.OWC10Api._Range EntireRow { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.Font Font { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        object Formula { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        object FormulaArray { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        object FormulaLocal { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        object HasArray { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        object HasFormula { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        object Height { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        bool Hidden { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        object HorizontalAlignment { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string HTMLData { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.Hyperlink Hyperlink { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.Interior Interior { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        object Left { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        object Locked { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        NetOffice.OWC10Api._Range MergeArea { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        object MergeCells { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.Name Name { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        NetOffice.OWC10Api._Range Next { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        object NumberFormat { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        new NetOffice.OWC10Api._Range Offset { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.Worksheet Parent { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        object PrefixCharacter { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        NetOffice.OWC10Api._Range Previous { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="cell1">object cell1</param>
        /// <param name="cell2">optional object cell2</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OWC10Api._Range get_Range(object cell1, object cell2);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_Range
        /// </summary>
        /// <param name="cell1">object cell1</param>
        /// <param name="cell2">optional object cell2</param>
        [SupportByVersion("OWC10", 1), Redirect("get_Range")]
        NetOffice.OWC10Api._Range Range(object cell1, object cell2);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="cell1">object cell1</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OWC10Api._Range get_Range(object cell1);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_Range
        /// </summary>
        /// <param name="cell1">object cell1</param>
        [SupportByVersion("OWC10", 1), Redirect("get_Range")]
        NetOffice.OWC10Api._Range Range(object cell1);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        object ReadingOrder { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        Int32 Row { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        object RowHeight { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        NetOffice.OWC10Api._Range Rows { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        object Text { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        object Top { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        object UseStandardHeight { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        object UseStandardWidth { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        new  object Value { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object Value2 { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        object VerticalAlignment { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        object Width { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.Worksheet Worksheet { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        void Activate();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="field">optional object field</param>
        /// <param name="criteria1">optional object criteria1</param>
        /// <param name="_operator">optional object operator</param>
        /// <param name="criteria2">optional object criteria2</param>
        /// <param name="visibleDropDown">optional object visibleDropDown</param>
        [SupportByVersion("OWC10", 1)]
        void AutoFilter(object field, object criteria1, object _operator, object criteria2, object visibleDropDown);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void AutoFilter();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="field">optional object field</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void AutoFilter(object field);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="field">optional object field</param>
        /// <param name="criteria1">optional object criteria1</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void AutoFilter(object field, object criteria1);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="field">optional object field</param>
        /// <param name="criteria1">optional object criteria1</param>
        /// <param name="_operator">optional object operator</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void AutoFilter(object field, object criteria1, object _operator);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="field">optional object field</param>
        /// <param name="criteria1">optional object criteria1</param>
        /// <param name="_operator">optional object operator</param>
        /// <param name="criteria2">optional object criteria2</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void AutoFilter(object field, object criteria1, object _operator, object criteria2);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        void AutoFit();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="lineStyle">optional object lineStyle</param>
        /// <param name="weight">optional NetOffice.OWC10Api.Enums.XlBorderWeight Weight = 2</param>
        /// <param name="colorIndex">optional NetOffice.OWC10Api.Enums.XlColorIndex ColorIndex = -4105</param>
        /// <param name="color">optional object color</param>
        [SupportByVersion("OWC10", 1)]
        void BorderAround(object lineStyle, object weight, object colorIndex, object color);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void BorderAround();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="lineStyle">optional object lineStyle</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void BorderAround(object lineStyle);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="lineStyle">optional object lineStyle</param>
        /// <param name="weight">optional NetOffice.OWC10Api.Enums.XlBorderWeight Weight = 2</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void BorderAround(object lineStyle, object weight);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="lineStyle">optional object lineStyle</param>
        /// <param name="weight">optional NetOffice.OWC10Api.Enums.XlBorderWeight Weight = 2</param>
        /// <param name="colorIndex">optional NetOffice.OWC10Api.Enums.XlColorIndex ColorIndex = -4105</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void BorderAround(object lineStyle, object weight, object colorIndex);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        void Calculate();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        void Clear();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        void ClearFormats();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        void ClearContents();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="destination">optional object destination</param>
        [SupportByVersion("OWC10", 1)]
        void Copy(object destination);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void Copy();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="data">object data</param>
        /// <param name="maxRows">optional object maxRows</param>
        /// <param name="maxColumns">optional object maxColumns</param>
        [SupportByVersion("OWC10", 1)]
        Int32 CopyFromRecordset(object data, object maxRows, object maxColumns);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="data">object data</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        Int32 CopyFromRecordset(object data);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="data">object data</param>
        /// <param name="maxRows">optional object maxRows</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        Int32 CopyFromRecordset(object data, object maxRows);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="destination">optional object destination</param>
        [SupportByVersion("OWC10", 1)]
        void Cut(object destination);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void Cut();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="shift">optional object shift</param>
        [SupportByVersion("OWC10", 1)]
        void Delete(object shift);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void Delete();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        void FillDown();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        void FillRight();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        /// <param name="lookIn">optional object lookIn</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        /// <param name="searchDirection">optional NetOffice.OWC10Api.Enums.XlSearchDirection SearchDirection = 1</param>
        /// <param name="matchCase">optional object matchCase</param>
        /// <param name="matchByte">optional object matchByte</param>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        NetOffice.OWC10Api._Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection, object matchCase, object matchByte);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="what">object what</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api._Range Find(object what);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api._Range Find(object what, object after);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        /// <param name="lookIn">optional object lookIn</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api._Range Find(object what, object after, object lookIn);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        /// <param name="lookIn">optional object lookIn</param>
        /// <param name="lookAt">optional object lookAt</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api._Range Find(object what, object after, object lookIn, object lookAt);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        /// <param name="lookIn">optional object lookIn</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api._Range Find(object what, object after, object lookIn, object lookAt, object searchOrder);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        /// <param name="lookIn">optional object lookIn</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        /// <param name="searchDirection">optional NetOffice.OWC10Api.Enums.XlSearchDirection SearchDirection = 1</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api._Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        /// <param name="lookIn">optional object lookIn</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        /// <param name="searchDirection">optional NetOffice.OWC10Api.Enums.XlSearchDirection SearchDirection = 1</param>
        /// <param name="matchCase">optional object matchCase</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api._Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection, object matchCase);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="after">optional object after</param>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        NetOffice.OWC10Api._Range FindNext(object after);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api._Range FindNext();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="after">optional object after</param>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        NetOffice.OWC10Api._Range FindPrevious(object after);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api._Range FindPrevious();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="shift">optional object shift</param>
        [SupportByVersion("OWC10", 1)]
        void Insert(object shift);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void Insert();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="file">string file</param>
        /// <param name="delimiters">optional string Delimiters = </param>
        /// <param name="consecutiveDelimAsOne">optional bool ConsecutiveDelimAsOne = false</param>
        /// <param name="textQualifier">optional string TextQualifier = \042</param>
        [SupportByVersion("OWC10", 1)]
        void LoadText(string file, object delimiters, object consecutiveDelimAsOne, object textQualifier);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="file">string file</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void LoadText(string file);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="file">string file</param>
        /// <param name="delimiters">optional string Delimiters = </param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void LoadText(string file, object delimiters);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="file">string file</param>
        /// <param name="delimiters">optional string Delimiters = </param>
        /// <param name="consecutiveDelimAsOne">optional bool ConsecutiveDelimAsOne = false</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void LoadText(string file, object delimiters, object consecutiveDelimAsOne);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="across">optional object across</param>
        [SupportByVersion("OWC10", 1)]
        void Merge(object across);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void Merge();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="text">string text</param>
        /// <param name="delimiters">optional string Delimiters = </param>
        /// <param name="consecutiveDelimAsOne">optional bool ConsecutiveDelimAsOne = false</param>
        /// <param name="textQualifier">optional string TextQualifier = \042</param>
        [SupportByVersion("OWC10", 1)]
        void ParseText(string text, object delimiters, object consecutiveDelimAsOne, object textQualifier);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="text">string text</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void ParseText(string text);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="text">string text</param>
        /// <param name="delimiters">optional string Delimiters = </param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void ParseText(string text, object delimiters);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="text">string text</param>
        /// <param name="delimiters">optional string Delimiters = </param>
        /// <param name="consecutiveDelimAsOne">optional bool ConsecutiveDelimAsOne = false</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void ParseText(string text, object delimiters, object consecutiveDelimAsOne);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        void Paste();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        void Select();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        void Show();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="columnKey">optional Int32 ColumnKey = 1</param>
        /// <param name="order">optional NetOffice.OWC10Api.Enums.XlSortOrder Order = 1</param>
        /// <param name="header">optional NetOffice.OWC10Api.Enums.XlYesNoGuess Header = 2</param>
        [SupportByVersion("OWC10", 1)]
        void Sort(object columnKey, object order, object header);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void Sort();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="columnKey">optional Int32 ColumnKey = 1</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void Sort(object columnKey);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="columnKey">optional Int32 ColumnKey = 1</param>
        /// <param name="order">optional NetOffice.OWC10Api.Enums.XlSortOrder Order = 1</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void Sort(object columnKey, object order);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        void UnMerge();

        #endregion

        #region IEnumerable<object>

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        new IEnumerator<object> GetEnumerator();

        #endregion
    }
}
