using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.ExcelApi
{
    /// <summary>
    /// Range
    /// </summary>
    [SyntaxBypass]
    public interface Range_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
        /// <param name="external">optional object external</param>
        /// <param name="relativeTo">optional object relativeTo</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837625.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external, object relativeTo);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Address
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837625.aspx </remarks>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
        /// <param name="external">optional object external</param>
        /// <param name="relativeTo">optional object relativeTo</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Address")]
        string Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external, object relativeTo);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837625.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_Address(object rowAbsolute);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Address
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837625.aspx </remarks>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Address")]
        string Address(object rowAbsolute);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837625.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_Address(object rowAbsolute, object columnAbsolute);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Address
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837625.aspx </remarks>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Address")]
        string Address(object rowAbsolute, object columnAbsolute);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837625.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_Address(object rowAbsolute, object columnAbsolute, object referenceStyle);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Address
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837625.aspx </remarks>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Address")]
        string Address(object rowAbsolute, object columnAbsolute, object referenceStyle);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
        /// <param name="external">optional object external</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837625.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Address
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837625.aspx </remarks>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
        /// <param name="external">optional object external</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Address")]
        string Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
        /// <param name="external">optional object external</param>
        /// <param name="relativeTo">optional object relativeTo</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194992.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_AddressLocal(object rowAbsolute, object columnAbsolute, object referenceStyle, object external, object relativeTo);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_AddressLocal
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194992.aspx </remarks>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
        /// <param name="external">optional object external</param>
        /// <param name="relativeTo">optional object relativeTo</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_AddressLocal")]
        string AddressLocal(object rowAbsolute, object columnAbsolute, object referenceStyle, object external, object relativeTo);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194992.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_AddressLocal(object rowAbsolute);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_AddressLocal
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194992.aspx </remarks>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_AddressLocal")]
        string AddressLocal(object rowAbsolute);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194992.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_AddressLocal(object rowAbsolute, object columnAbsolute);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_AddressLocal
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194992.aspx </remarks>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_AddressLocal")]
        string AddressLocal(object rowAbsolute, object columnAbsolute);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194992.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_AddressLocal(object rowAbsolute, object columnAbsolute, object referenceStyle);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_AddressLocal
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194992.aspx </remarks>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_AddressLocal")]
        string AddressLocal(object rowAbsolute, object columnAbsolute, object referenceStyle);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
        /// <param name="external">optional object external</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194992.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_AddressLocal(object rowAbsolute, object columnAbsolute, object referenceStyle, object external);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_AddressLocal
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194992.aspx </remarks>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
        /// <param name="external">optional object external</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_AddressLocal")]
        string AddressLocal(object rowAbsolute, object columnAbsolute, object referenceStyle, object external);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional object start</param>
        /// <param name="length">optional object length</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198232.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.ExcelApi.Characters get_Characters(object start, object length);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198232.aspx </remarks>
        /// <param name="start">optional object start</param>
        /// <param name="length">optional object length</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Characters")]
        NetOffice.ExcelApi.Characters Characters(object start, object length);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional object start</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198232.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.ExcelApi.Characters get_Characters(object start);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198232.aspx </remarks>
        /// <param name="start">optional object start</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Characters")]
        NetOffice.ExcelApi.Characters Characters(object start);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="rowOffset">optional object rowOffset</param>
        /// <param name="columnOffset">optional object columnOffset</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840060.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.ExcelApi.Range get_Offset(object rowOffset, object columnOffset);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Offset
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840060.aspx </remarks>
        /// <param name="rowOffset">optional object rowOffset</param>
        /// <param name="columnOffset">optional object columnOffset</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Offset")]
        NetOffice.ExcelApi.Range Offset(object rowOffset, object columnOffset);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="rowOffset">optional object rowOffset</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840060.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.ExcelApi.Range get_Offset(object rowOffset);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Offset
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840060.aspx </remarks>
        /// <param name="rowOffset">optional object rowOffset</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Offset")]
        NetOffice.ExcelApi.Range Offset(object rowOffset);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="cell1">object cell1</param>
        /// <param name="cell2">optional object cell2</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834676.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.ExcelApi.Range get_Range(object cell1, object cell2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Range
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834676.aspx </remarks>
        /// <param name="cell1">object cell1</param>
        /// <param name="cell2">optional object cell2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Range")]
        NetOffice.ExcelApi.Range Range(object cell1, object cell2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="cell1">object cell1</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834676.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.ExcelApi.Range get_Range(object cell1)
            ;
        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Range
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834676.aspx </remarks>
        /// <param name="cell1">object cell1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Range")]
        NetOffice.ExcelApi.Range Range(object cell1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="rowSize">optional object rowSize</param>
        /// <param name="columnSize">optional object columnSize</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193274.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.ExcelApi.Range get_Resize(object rowSize, object columnSize);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Resize
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193274.aspx </remarks>
        /// <param name="rowSize">optional object rowSize</param>
        /// <param name="columnSize">optional object columnSize</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Resize")]
        NetOffice.ExcelApi.Range Resize(object rowSize, object columnSize);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="rowSize">optional object rowSize</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193274.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.ExcelApi.Range get_Resize(object rowSize);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Resize
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193274.aspx </remarks>
        /// <param name="rowSize">optional object rowSize</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Resize")]
        NetOffice.ExcelApi.Range Resize(object rowSize);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="rangeValueDataType">optional object rangeValueDataType</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195193.aspx
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_Value(object rangeValueDataType);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="rangeValueDataType">optional object rangeValueDataType</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_Value(object rangeValueDataType, object value);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Alias for get_Value
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195193.aspx </remarks>
        /// <param name="rangeValueDataType">optional object rangeValueDataType</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16), Redirect("get_Value")]
        object Value(object rangeValueDataType);

        #endregion
    }

    /// <summary>
    /// DispatchInterface Range 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838238.aspx </remarks>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Excel", 9, 10, 11, 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "_Default")]
	[TypeId("00020846-0000-0000-C000-000000000046")]
    public interface Range : Range_, IEnumerableProvider<NetOffice.ExcelApi.Range>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193954.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839656.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196885.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197727.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AddIndent { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837625.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        new string Address { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194992.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        new string AddressLocal { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196243.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Areas Areas { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822605.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Borders Borders { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196273.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range Cells { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198232.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        new NetOffice.ExcelApi.Characters Characters { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198200.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Column { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837125.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range Columns { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837430.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object ColumnWidth { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193349.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Count { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194268.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range CurrentArray { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196678.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range CurrentRegion { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// Custom Indexer
        /// </summary>
        /// <param name="rowIndex">optional object rowIndex</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty, CustomIndexer]
        NetOffice.ExcelApi.Range this[object rowIndex] { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="rowIndex">optional object rowIndex</param>
        /// <param name="columnIndex">optional object columnIndex</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.ExcelApi.Range this[object rowIndex, object columnIndex] { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197707.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range Dependents { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195386.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range DirectDependents { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839667.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range DirectPrecedents { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839539.aspx </remarks>
        /// <param name="direction">NetOffice.ExcelApi.Enums.XlDirection direction</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.ExcelApi.Range get_End(NetOffice.ExcelApi.Enums.XlDirection direction);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_End
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839539.aspx </remarks>
        /// <param name="direction">NetOffice.ExcelApi.Enums.XlDirection direction</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_End")]
        NetOffice.ExcelApi.Range End(NetOffice.ExcelApi.Enums.XlDirection direction);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834476.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range EntireColumn { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836836.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range EntireRow { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839770.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Font Font { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838835.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Formula { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837104.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object FormulaArray { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlFormulaLabel FormulaLabel { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838209.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object FormulaHidden { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838851.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object FormulaLocal { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823188.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object FormulaR1C1 { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838568.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object FormulaR1C1Local { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841166.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object HasArray { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837123.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object HasFormula { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840142.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Height { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834657.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Hidden { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822120.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object HorizontalAlignment { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840976.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object IndentLevel { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836210.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Interior Interior { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821859.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Left { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839644.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 ListHeaderRows { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834387.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlLocationInTable LocationInTable { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836172.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Locked { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822300.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range MergeArea { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197310.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object MergeCells { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196817.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Name { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193940.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range Next { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196401.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object NumberFormat { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840206.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object NumberFormatLocal { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840060.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        new NetOffice.ExcelApi.Range Offset { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198186.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Orientation { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838455.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object OutlineLevel { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193644.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 PageBreak { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820916.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.PivotField PivotField { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193001.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.PivotItem PivotItem { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837826.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.PivotTable PivotTable { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196936.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range Precedents { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194949.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object PrefixCharacter { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822759.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range Previous { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821864.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.QueryTable QueryTable { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193274.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        new NetOffice.ExcelApi.Range Resize { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196952.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Row { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193926.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object RowHeight { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195745.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range Rows { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194526.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object ShowDetail { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841203.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object ShrinkToFit { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193277.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.SoundNote SoundNote { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834326.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Style { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841152.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Summary { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840217.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Text { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193765.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Top { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821216.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object UseStandardHeight { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836425.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object UseStandardWidth { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839408.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Validation Validation { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195193.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        new object Value { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193553.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Value2 { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837977.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object VerticalAlignment { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823126.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Width { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837845.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Worksheet Worksheet { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821514.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object WrapText { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836191.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Comment Comment { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836776.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Phonetic Phonetic { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822167.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.FormatConditions FormatConditions { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840917.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 ReadingOrder { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839653.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Hyperlinks Hyperlinks { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841233.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Phonetics Phonetics { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193914.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string ID { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836438.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.PivotCell PivotCell { get; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835287.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Errors Errors { get; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.SmartTags SmartTags { get; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837052.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        bool AllowEdit { get; }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838413.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.ListObject ListObject { get; }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835889.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.XPath XPath { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840072.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.ExcelApi.Actions ServerActions { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822495.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string MDX { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196838.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        object CountLarge { get; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822140.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        NetOffice.ExcelApi.SparklineGroups SparklineGroups { get; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838814.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        NetOffice.ExcelApi.DisplayFormat DisplayFormat { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837085.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Activate();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841242.aspx </remarks>
        /// <param name="action">NetOffice.ExcelApi.Enums.XlFilterAction action</param>
        /// <param name="criteriaRange">optional object criteriaRange</param>
        /// <param name="copyToRange">optional object copyToRange</param>
        /// <param name="unique">optional object unique</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AdvancedFilter(NetOffice.ExcelApi.Enums.XlFilterAction action, object criteriaRange, object copyToRange, object unique);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841242.aspx </remarks>
        /// <param name="action">NetOffice.ExcelApi.Enums.XlFilterAction action</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AdvancedFilter(NetOffice.ExcelApi.Enums.XlFilterAction action);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841242.aspx </remarks>
        /// <param name="action">NetOffice.ExcelApi.Enums.XlFilterAction action</param>
        /// <param name="criteriaRange">optional object criteriaRange</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AdvancedFilter(NetOffice.ExcelApi.Enums.XlFilterAction action, object criteriaRange);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841242.aspx </remarks>
        /// <param name="action">NetOffice.ExcelApi.Enums.XlFilterAction action</param>
        /// <param name="criteriaRange">optional object criteriaRange</param>
        /// <param name="copyToRange">optional object copyToRange</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AdvancedFilter(NetOffice.ExcelApi.Enums.XlFilterAction action, object criteriaRange, object copyToRange);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196578.aspx </remarks>
        /// <param name="names">optional object names</param>
        /// <param name="ignoreRelativeAbsolute">optional object ignoreRelativeAbsolute</param>
        /// <param name="useRowColumnNames">optional object useRowColumnNames</param>
        /// <param name="omitColumn">optional object omitColumn</param>
        /// <param name="omitRow">optional object omitRow</param>
        /// <param name="order">optional NetOffice.ExcelApi.Enums.XlApplyNamesOrder Order = 1</param>
        /// <param name="appendLast">optional object appendLast</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object ApplyNames(object names, object ignoreRelativeAbsolute, object useRowColumnNames, object omitColumn, object omitRow, object order, object appendLast);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196578.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object ApplyNames();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196578.aspx </remarks>
        /// <param name="names">optional object names</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object ApplyNames(object names);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196578.aspx </remarks>
        /// <param name="names">optional object names</param>
        /// <param name="ignoreRelativeAbsolute">optional object ignoreRelativeAbsolute</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object ApplyNames(object names, object ignoreRelativeAbsolute);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196578.aspx </remarks>
        /// <param name="names">optional object names</param>
        /// <param name="ignoreRelativeAbsolute">optional object ignoreRelativeAbsolute</param>
        /// <param name="useRowColumnNames">optional object useRowColumnNames</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object ApplyNames(object names, object ignoreRelativeAbsolute, object useRowColumnNames);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196578.aspx </remarks>
        /// <param name="names">optional object names</param>
        /// <param name="ignoreRelativeAbsolute">optional object ignoreRelativeAbsolute</param>
        /// <param name="useRowColumnNames">optional object useRowColumnNames</param>
        /// <param name="omitColumn">optional object omitColumn</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object ApplyNames(object names, object ignoreRelativeAbsolute, object useRowColumnNames, object omitColumn);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196578.aspx </remarks>
        /// <param name="names">optional object names</param>
        /// <param name="ignoreRelativeAbsolute">optional object ignoreRelativeAbsolute</param>
        /// <param name="useRowColumnNames">optional object useRowColumnNames</param>
        /// <param name="omitColumn">optional object omitColumn</param>
        /// <param name="omitRow">optional object omitRow</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object ApplyNames(object names, object ignoreRelativeAbsolute, object useRowColumnNames, object omitColumn, object omitRow);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196578.aspx </remarks>
        /// <param name="names">optional object names</param>
        /// <param name="ignoreRelativeAbsolute">optional object ignoreRelativeAbsolute</param>
        /// <param name="useRowColumnNames">optional object useRowColumnNames</param>
        /// <param name="omitColumn">optional object omitColumn</param>
        /// <param name="omitRow">optional object omitRow</param>
        /// <param name="order">optional NetOffice.ExcelApi.Enums.XlApplyNamesOrder Order = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object ApplyNames(object names, object ignoreRelativeAbsolute, object useRowColumnNames, object omitColumn, object omitRow, object order);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840485.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object ApplyOutlineStyles();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822899.aspx </remarks>
        /// <param name="_string">string string</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string AutoComplete(string _string);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195345.aspx </remarks>
        /// <param name="destination">NetOffice.ExcelApi.Range destination</param>
        /// <param name="type">optional NetOffice.ExcelApi.Enums.XlAutoFillType Type = 0</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AutoFill(NetOffice.ExcelApi.Range destination, object type);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195345.aspx </remarks>
        /// <param name="destination">NetOffice.ExcelApi.Range destination</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AutoFill(NetOffice.ExcelApi.Range destination);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193884.aspx </remarks>
        /// <param name="field">optional object field</param>
        /// <param name="criteria1">optional object criteria1</param>
        /// <param name="_operator">optional NetOffice.ExcelApi.Enums.XlAutoFilterOperator Operator = 1</param>
        /// <param name="criteria2">optional object criteria2</param>
        /// <param name="visibleDropDown">optional object visibleDropDown</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AutoFilter(object field, object criteria1, object _operator, object criteria2, object visibleDropDown);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193884.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AutoFilter();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193884.aspx </remarks>
        /// <param name="field">optional object field</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AutoFilter(object field);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193884.aspx </remarks>
        /// <param name="field">optional object field</param>
        /// <param name="criteria1">optional object criteria1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AutoFilter(object field, object criteria1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193884.aspx </remarks>
        /// <param name="field">optional object field</param>
        /// <param name="criteria1">optional object criteria1</param>
        /// <param name="_operator">optional NetOffice.ExcelApi.Enums.XlAutoFilterOperator Operator = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AutoFilter(object field, object criteria1, object _operator);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193884.aspx </remarks>
        /// <param name="field">optional object field</param>
        /// <param name="criteria1">optional object criteria1</param>
        /// <param name="_operator">optional NetOffice.ExcelApi.Enums.XlAutoFilterOperator Operator = 1</param>
        /// <param name="criteria2">optional object criteria2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AutoFilter(object field, object criteria1, object _operator, object criteria2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820840.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AutoFit();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">optional NetOffice.ExcelApi.Enums.XlRangeAutoFormat Format = 1</param>
        /// <param name="number">optional object number</param>
        /// <param name="font">optional object font</param>
        /// <param name="alignment">optional object alignment</param>
        /// <param name="border">optional object border</param>
        /// <param name="pattern">optional object pattern</param>
        /// <param name="width">optional object width</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AutoFormat(object format, object number, object font, object alignment, object border, object pattern, object width);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AutoFormat();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">optional NetOffice.ExcelApi.Enums.XlRangeAutoFormat Format = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AutoFormat(object format);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">optional NetOffice.ExcelApi.Enums.XlRangeAutoFormat Format = 1</param>
        /// <param name="number">optional object number</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AutoFormat(object format, object number);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">optional NetOffice.ExcelApi.Enums.XlRangeAutoFormat Format = 1</param>
        /// <param name="number">optional object number</param>
        /// <param name="font">optional object font</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AutoFormat(object format, object number, object font);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">optional NetOffice.ExcelApi.Enums.XlRangeAutoFormat Format = 1</param>
        /// <param name="number">optional object number</param>
        /// <param name="font">optional object font</param>
        /// <param name="alignment">optional object alignment</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AutoFormat(object format, object number, object font, object alignment);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">optional NetOffice.ExcelApi.Enums.XlRangeAutoFormat Format = 1</param>
        /// <param name="number">optional object number</param>
        /// <param name="font">optional object font</param>
        /// <param name="alignment">optional object alignment</param>
        /// <param name="border">optional object border</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AutoFormat(object format, object number, object font, object alignment, object border);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">optional NetOffice.ExcelApi.Enums.XlRangeAutoFormat Format = 1</param>
        /// <param name="number">optional object number</param>
        /// <param name="font">optional object font</param>
        /// <param name="alignment">optional object alignment</param>
        /// <param name="border">optional object border</param>
        /// <param name="pattern">optional object pattern</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AutoFormat(object format, object number, object font, object alignment, object border, object pattern);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837148.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AutoOutline();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197210.aspx </remarks>
        /// <param name="lineStyle">optional object lineStyle</param>
        /// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
        /// <param name="colorIndex">optional NetOffice.ExcelApi.Enums.XlColorIndex ColorIndex = -4105</param>
        /// <param name="color">optional object color</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object BorderAround(object lineStyle, object weight, object colorIndex, object color);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197210.aspx </remarks>
        /// <param name="lineStyle">optional object lineStyle</param>
        /// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
        /// <param name="colorIndex">optional NetOffice.ExcelApi.Enums.XlColorIndex ColorIndex = -4105</param>
        /// <param name="color">optional object color</param>
        /// <param name="themeColor">optional object themeColor</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        object BorderAround(object lineStyle, object weight, object colorIndex, object color, object themeColor);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197210.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object BorderAround();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197210.aspx </remarks>
        /// <param name="lineStyle">optional object lineStyle</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object BorderAround(object lineStyle);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197210.aspx </remarks>
        /// <param name="lineStyle">optional object lineStyle</param>
        /// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object BorderAround(object lineStyle, object weight);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197210.aspx </remarks>
        /// <param name="lineStyle">optional object lineStyle</param>
        /// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
        /// <param name="colorIndex">optional NetOffice.ExcelApi.Enums.XlColorIndex ColorIndex = -4105</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object BorderAround(object lineStyle, object weight, object colorIndex);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834613.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Calculate();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195092.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        /// <param name="spellLang">optional object spellLang</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object spellLang);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195092.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object CheckSpelling();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195092.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object CheckSpelling(object customDictionary);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195092.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object CheckSpelling(object customDictionary, object ignoreUppercase);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195092.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820947.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Clear();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835589.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object ClearContents();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196576.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object ClearFormats();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195197.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object ClearNotes();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834886.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object ClearOutline();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197752.aspx </remarks>
        /// <param name="comparison">object comparison</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range ColumnDifferences(object comparison);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839594.aspx </remarks>
        /// <param name="sources">optional object sources</param>
        /// <param name="function">optional object function</param>
        /// <param name="topRow">optional object topRow</param>
        /// <param name="leftColumn">optional object leftColumn</param>
        /// <param name="createLinks">optional object createLinks</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Consolidate(object sources, object function, object topRow, object leftColumn, object createLinks);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839594.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Consolidate();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839594.aspx </remarks>
        /// <param name="sources">optional object sources</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Consolidate(object sources);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839594.aspx </remarks>
        /// <param name="sources">optional object sources</param>
        /// <param name="function">optional object function</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Consolidate(object sources, object function);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839594.aspx </remarks>
        /// <param name="sources">optional object sources</param>
        /// <param name="function">optional object function</param>
        /// <param name="topRow">optional object topRow</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Consolidate(object sources, object function, object topRow);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839594.aspx </remarks>
        /// <param name="sources">optional object sources</param>
        /// <param name="function">optional object function</param>
        /// <param name="topRow">optional object topRow</param>
        /// <param name="leftColumn">optional object leftColumn</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Consolidate(object sources, object function, object topRow, object leftColumn);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837760.aspx </remarks>
        /// <param name="destination">optional object destination</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Copy(object destination);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837760.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Copy();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839240.aspx </remarks>
        /// <param name="data">object data</param>
        /// <param name="maxRows">optional object maxRows</param>
        /// <param name="maxColumns">optional object maxColumns</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 CopyFromRecordset(object data, object maxRows, object maxColumns);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839240.aspx </remarks>
        /// <param name="data">object data</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 CopyFromRecordset(object data);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839240.aspx </remarks>
        /// <param name="data">object data</param>
        /// <param name="maxRows">optional object maxRows</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 CopyFromRecordset(object data, object maxRows);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193594.aspx </remarks>
        /// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
        /// <param name="format">optional NetOffice.ExcelApi.Enums.XlCopyPictureFormat Format = -4147</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object CopyPicture(object appearance, object format);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193594.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object CopyPicture();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193594.aspx </remarks>
        /// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object CopyPicture(object appearance);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192960.aspx </remarks>
        /// <param name="top">optional object top</param>
        /// <param name="left">optional object left</param>
        /// <param name="bottom">optional object bottom</param>
        /// <param name="right">optional object right</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object CreateNames(object top, object left, object bottom, object right);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192960.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object CreateNames();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192960.aspx </remarks>
        /// <param name="top">optional object top</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object CreateNames(object top);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192960.aspx </remarks>
        /// <param name="top">optional object top</param>
        /// <param name="left">optional object left</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object CreateNames(object top, object left);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192960.aspx </remarks>
        /// <param name="top">optional object top</param>
        /// <param name="left">optional object left</param>
        /// <param name="bottom">optional object bottom</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object CreateNames(object top, object left, object bottom);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="edition">optional object edition</param>
        /// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
        /// <param name="containsPICT">optional object containsPICT</param>
        /// <param name="containsBIFF">optional object containsBIFF</param>
        /// <param name="containsRTF">optional object containsRTF</param>
        /// <param name="containsVALU">optional object containsVALU</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object CreatePublisher(object edition, object appearance, object containsPICT, object containsBIFF, object containsRTF, object containsVALU);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object CreatePublisher();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="edition">optional object edition</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object CreatePublisher(object edition);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="edition">optional object edition</param>
        /// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object CreatePublisher(object edition, object appearance);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="edition">optional object edition</param>
        /// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
        /// <param name="containsPICT">optional object containsPICT</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object CreatePublisher(object edition, object appearance, object containsPICT);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="edition">optional object edition</param>
        /// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
        /// <param name="containsPICT">optional object containsPICT</param>
        /// <param name="containsBIFF">optional object containsBIFF</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object CreatePublisher(object edition, object appearance, object containsPICT, object containsBIFF);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="edition">optional object edition</param>
        /// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
        /// <param name="containsPICT">optional object containsPICT</param>
        /// <param name="containsBIFF">optional object containsBIFF</param>
        /// <param name="containsRTF">optional object containsRTF</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object CreatePublisher(object edition, object appearance, object containsPICT, object containsBIFF, object containsRTF);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838374.aspx </remarks>
        /// <param name="destination">optional object destination</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Cut(object destination);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838374.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Cut();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839262.aspx </remarks>
        /// <param name="rowcol">optional object rowcol</param>
        /// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataSeriesType Type = -4132</param>
        /// <param name="date">optional NetOffice.ExcelApi.Enums.XlDataSeriesDate Date = 1</param>
        /// <param name="step">optional object step</param>
        /// <param name="stop">optional object stop</param>
        /// <param name="trend">optional object trend</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object DataSeries(object rowcol, object type, object date, object step, object stop, object trend);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839262.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object DataSeries();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839262.aspx </remarks>
        /// <param name="rowcol">optional object rowcol</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object DataSeries(object rowcol);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839262.aspx </remarks>
        /// <param name="rowcol">optional object rowcol</param>
        /// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataSeriesType Type = -4132</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object DataSeries(object rowcol, object type);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839262.aspx </remarks>
        /// <param name="rowcol">optional object rowcol</param>
        /// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataSeriesType Type = -4132</param>
        /// <param name="date">optional NetOffice.ExcelApi.Enums.XlDataSeriesDate Date = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object DataSeries(object rowcol, object type, object date);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839262.aspx </remarks>
        /// <param name="rowcol">optional object rowcol</param>
        /// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataSeriesType Type = -4132</param>
        /// <param name="date">optional NetOffice.ExcelApi.Enums.XlDataSeriesDate Date = 1</param>
        /// <param name="step">optional object step</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object DataSeries(object rowcol, object type, object date, object step);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839262.aspx </remarks>
        /// <param name="rowcol">optional object rowcol</param>
        /// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataSeriesType Type = -4132</param>
        /// <param name="date">optional NetOffice.ExcelApi.Enums.XlDataSeriesDate Date = 1</param>
        /// <param name="step">optional object step</param>
        /// <param name="stop">optional object stop</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object DataSeries(object rowcol, object type, object date, object step, object stop);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834641.aspx </remarks>
        /// <param name="shift">optional object shift</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Delete(object shift);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834641.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Delete();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839446.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object DialogBox();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821204.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlEditionType type</param>
        /// <param name="option">NetOffice.ExcelApi.Enums.XlEditionOptionsOption option</param>
        /// <param name="name">optional object name</param>
        /// <param name="reference">optional object reference</param>
        /// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
        /// <param name="chartSize">optional NetOffice.ExcelApi.Enums.XlPictureAppearance ChartSize = 1</param>
        /// <param name="format">optional object format</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object EditionOptions(NetOffice.ExcelApi.Enums.XlEditionType type, NetOffice.ExcelApi.Enums.XlEditionOptionsOption option, object name, object reference, object appearance, object chartSize, object format);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821204.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlEditionType type</param>
        /// <param name="option">NetOffice.ExcelApi.Enums.XlEditionOptionsOption option</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object EditionOptions(NetOffice.ExcelApi.Enums.XlEditionType type, NetOffice.ExcelApi.Enums.XlEditionOptionsOption option);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821204.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlEditionType type</param>
        /// <param name="option">NetOffice.ExcelApi.Enums.XlEditionOptionsOption option</param>
        /// <param name="name">optional object name</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object EditionOptions(NetOffice.ExcelApi.Enums.XlEditionType type, NetOffice.ExcelApi.Enums.XlEditionOptionsOption option, object name);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821204.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlEditionType type</param>
        /// <param name="option">NetOffice.ExcelApi.Enums.XlEditionOptionsOption option</param>
        /// <param name="name">optional object name</param>
        /// <param name="reference">optional object reference</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object EditionOptions(NetOffice.ExcelApi.Enums.XlEditionType type, NetOffice.ExcelApi.Enums.XlEditionOptionsOption option, object name, object reference);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821204.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlEditionType type</param>
        /// <param name="option">NetOffice.ExcelApi.Enums.XlEditionOptionsOption option</param>
        /// <param name="name">optional object name</param>
        /// <param name="reference">optional object reference</param>
        /// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object EditionOptions(NetOffice.ExcelApi.Enums.XlEditionType type, NetOffice.ExcelApi.Enums.XlEditionOptionsOption option, object name, object reference, object appearance);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821204.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlEditionType type</param>
        /// <param name="option">NetOffice.ExcelApi.Enums.XlEditionOptionsOption option</param>
        /// <param name="name">optional object name</param>
        /// <param name="reference">optional object reference</param>
        /// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
        /// <param name="chartSize">optional NetOffice.ExcelApi.Enums.XlPictureAppearance ChartSize = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object EditionOptions(NetOffice.ExcelApi.Enums.XlEditionType type, NetOffice.ExcelApi.Enums.XlEditionOptionsOption option, object name, object reference, object appearance, object chartSize);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838404.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object FillDown();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197298.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object FillLeft();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837984.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object FillRight();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820752.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object FillUp();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839746.aspx </remarks>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        /// <param name="lookIn">optional object lookIn</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        /// <param name="searchDirection">optional NetOffice.ExcelApi.Enums.XlSearchDirection SearchDirection = 1</param>
        /// <param name="matchCase">optional object matchCase</param>
        /// <param name="matchByte">optional object matchByte</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection, object matchCase, object matchByte);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839746.aspx </remarks>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        /// <param name="lookIn">optional object lookIn</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        /// <param name="searchDirection">optional NetOffice.ExcelApi.Enums.XlSearchDirection SearchDirection = 1</param>
        /// <param name="matchCase">optional object matchCase</param>
        /// <param name="matchByte">optional object matchByte</param>
        /// <param name="searchFormat">optional object searchFormat</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection, object matchCase, object matchByte, object searchFormat);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839746.aspx </remarks>
        /// <param name="what">object what</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range Find(object what);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839746.aspx </remarks>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range Find(object what, object after);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839746.aspx </remarks>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        /// <param name="lookIn">optional object lookIn</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range Find(object what, object after, object lookIn);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839746.aspx </remarks>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        /// <param name="lookIn">optional object lookIn</param>
        /// <param name="lookAt">optional object lookAt</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range Find(object what, object after, object lookIn, object lookAt);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839746.aspx </remarks>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        /// <param name="lookIn">optional object lookIn</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range Find(object what, object after, object lookIn, object lookAt, object searchOrder);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839746.aspx </remarks>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        /// <param name="lookIn">optional object lookIn</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        /// <param name="searchDirection">optional NetOffice.ExcelApi.Enums.XlSearchDirection SearchDirection = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839746.aspx </remarks>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        /// <param name="lookIn">optional object lookIn</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        /// <param name="searchDirection">optional NetOffice.ExcelApi.Enums.XlSearchDirection SearchDirection = 1</param>
        /// <param name="matchCase">optional object matchCase</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection, object matchCase);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196143.aspx </remarks>
        /// <param name="after">optional object after</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range FindNext(object after);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196143.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range FindNext();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838614.aspx </remarks>
        /// <param name="after">optional object after</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range FindPrevious(object after);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838614.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range FindPrevious();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837608.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object FunctionWizard();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="goal">object goal</param>
        /// <param name="changingCell">NetOffice.ExcelApi.Range changingCell</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool GoalSeek(object goal, NetOffice.ExcelApi.Range changingCell);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839808.aspx </remarks>
        /// <param name="start">optional object start</param>
        /// <param name="end">optional object end</param>
        /// <param name="by">optional object by</param>
        /// <param name="periods">optional object periods</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Group(object start, object end, object by, object periods);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839808.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Group();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839808.aspx </remarks>
        /// <param name="start">optional object start</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Group(object start);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839808.aspx </remarks>
        /// <param name="start">optional object start</param>
        /// <param name="end">optional object end</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Group(object start, object end);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839808.aspx </remarks>
        /// <param name="start">optional object start</param>
        /// <param name="end">optional object end</param>
        /// <param name="by">optional object by</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Group(object start, object end, object by);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194874.aspx </remarks>
        /// <param name="insertAmount">Int32 insertAmount</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void InsertIndent(Int32 insertAmount);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840310.aspx </remarks>
        /// <param name="shift">optional object shift</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Insert(object shift);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840310.aspx </remarks>
        /// <param name="shift">optional object shift</param>
        /// <param name="copyOrigin">optional object copyOrigin</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        object Insert(object shift, object copyOrigin);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840310.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Insert();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841125.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Justify();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193257.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object ListNames();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840729.aspx </remarks>
        /// <param name="across">optional object across</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void Merge(object across);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840729.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void Merge();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840063.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void UnMerge();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822883.aspx </remarks>
        /// <param name="towardPrecedent">optional object towardPrecedent</param>
        /// <param name="arrowNumber">optional object arrowNumber</param>
        /// <param name="linkNumber">optional object linkNumber</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object NavigateArrow(object towardPrecedent, object arrowNumber, object linkNumber);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822883.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object NavigateArrow();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822883.aspx </remarks>
        /// <param name="towardPrecedent">optional object towardPrecedent</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object NavigateArrow(object towardPrecedent);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822883.aspx </remarks>
        /// <param name="towardPrecedent">optional object towardPrecedent</param>
        /// <param name="arrowNumber">optional object arrowNumber</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object NavigateArrow(object towardPrecedent, object arrowNumber);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839198.aspx </remarks>
        /// <param name="text">optional object text</param>
        /// <param name="start">optional object start</param>
        /// <param name="length">optional object length</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string NoteText(object text, object start, object length);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839198.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string NoteText();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839198.aspx </remarks>
        /// <param name="text">optional object text</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string NoteText(object text);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839198.aspx </remarks>
        /// <param name="text">optional object text</param>
        /// <param name="start">optional object start</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string NoteText(object text, object start);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196436.aspx </remarks>
        /// <param name="parseLine">optional object parseLine</param>
        /// <param name="destination">optional object destination</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Parse(object parseLine, object destination);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196436.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Parse();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196436.aspx </remarks>
        /// <param name="parseLine">optional object parseLine</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Parse(object parseLine);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839476.aspx </remarks>
        /// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
        /// <param name="operation">optional NetOffice.ExcelApi.Enums.XlPasteSpecialOperation Operation = -4142</param>
        /// <param name="skipBlanks">optional object skipBlanks</param>
        /// <param name="transpose">optional object transpose</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object PasteSpecial(object paste, object operation, object skipBlanks, object transpose);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839476.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object PasteSpecial();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839476.aspx </remarks>
        /// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object PasteSpecial(object paste);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839476.aspx </remarks>
        /// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
        /// <param name="operation">optional NetOffice.ExcelApi.Enums.XlPasteSpecialOperation Operation = -4142</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object PasteSpecial(object paste, object operation);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839476.aspx </remarks>
        /// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
        /// <param name="operation">optional NetOffice.ExcelApi.Enums.XlPasteSpecialOperation Operation = -4142</param>
        /// <param name="skipBlanks">optional object skipBlanks</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object PasteSpecial(object paste, object operation, object skipBlanks);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="prToFileName">optional object prToFileName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        object _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object _PrintOut();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object _PrintOut(object from);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object _PrintOut(object from, object to);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object _PrintOut(object from, object to, object copies);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object _PrintOut(object from, object to, object copies, object preview);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object _PrintOut(object from, object to, object copies, object preview, object activePrinter);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838066.aspx </remarks>
        /// <param name="enableChanges">optional object enableChanges</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object PrintPreview(object enableChanges);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838066.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object PrintPreview();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840556.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object RemoveSubtotal();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194086.aspx </remarks>
        /// <param name="what">object what</param>
        /// <param name="replacement">object replacement</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        /// <param name="matchCase">optional object matchCase</param>
        /// <param name="matchByte">optional object matchByte</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool Replace(object what, object replacement, object lookAt, object searchOrder, object matchCase, object matchByte);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194086.aspx </remarks>
        /// <param name="what">object what</param>
        /// <param name="replacement">object replacement</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        /// <param name="matchCase">optional object matchCase</param>
        /// <param name="matchByte">optional object matchByte</param>
        /// <param name="searchFormat">optional object searchFormat</param>
        /// <param name="replaceFormat">optional object replaceFormat</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        bool Replace(object what, object replacement, object lookAt, object searchOrder, object matchCase, object matchByte, object searchFormat, object replaceFormat);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194086.aspx </remarks>
        /// <param name="what">object what</param>
        /// <param name="replacement">object replacement</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool Replace(object what, object replacement);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194086.aspx </remarks>
        /// <param name="what">object what</param>
        /// <param name="replacement">object replacement</param>
        /// <param name="lookAt">optional object lookAt</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool Replace(object what, object replacement, object lookAt);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194086.aspx </remarks>
        /// <param name="what">object what</param>
        /// <param name="replacement">object replacement</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool Replace(object what, object replacement, object lookAt, object searchOrder);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194086.aspx </remarks>
        /// <param name="what">object what</param>
        /// <param name="replacement">object replacement</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        /// <param name="matchCase">optional object matchCase</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool Replace(object what, object replacement, object lookAt, object searchOrder, object matchCase);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194086.aspx </remarks>
        /// <param name="what">object what</param>
        /// <param name="replacement">object replacement</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        /// <param name="matchCase">optional object matchCase</param>
        /// <param name="matchByte">optional object matchByte</param>
        /// <param name="searchFormat">optional object searchFormat</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        bool Replace(object what, object replacement, object lookAt, object searchOrder, object matchCase, object matchByte, object searchFormat);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835292.aspx </remarks>
        /// <param name="comparison">object comparison</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range RowDifferences(object comparison);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        /// <param name="arg23">optional object arg23</param>
        /// <param name="arg24">optional object arg24</param>
        /// <param name="arg25">optional object arg25</param>
        /// <param name="arg26">optional object arg26</param>
        /// <param name="arg27">optional object arg27</param>
        /// <param name="arg28">optional object arg28</param>
        /// <param name="arg29">optional object arg29</param>
        /// <param name="arg30">optional object arg30</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29, object arg30);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3, object arg4, object arg5);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        /// <param name="arg23">optional object arg23</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        /// <param name="arg23">optional object arg23</param>
        /// <param name="arg24">optional object arg24</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        /// <param name="arg23">optional object arg23</param>
        /// <param name="arg24">optional object arg24</param>
        /// <param name="arg25">optional object arg25</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        /// <param name="arg23">optional object arg23</param>
        /// <param name="arg24">optional object arg24</param>
        /// <param name="arg25">optional object arg25</param>
        /// <param name="arg26">optional object arg26</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        /// <param name="arg23">optional object arg23</param>
        /// <param name="arg24">optional object arg24</param>
        /// <param name="arg25">optional object arg25</param>
        /// <param name="arg26">optional object arg26</param>
        /// <param name="arg27">optional object arg27</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        /// <param name="arg23">optional object arg23</param>
        /// <param name="arg24">optional object arg24</param>
        /// <param name="arg25">optional object arg25</param>
        /// <param name="arg26">optional object arg26</param>
        /// <param name="arg27">optional object arg27</param>
        /// <param name="arg28">optional object arg28</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        /// <param name="arg23">optional object arg23</param>
        /// <param name="arg24">optional object arg24</param>
        /// <param name="arg25">optional object arg25</param>
        /// <param name="arg26">optional object arg26</param>
        /// <param name="arg27">optional object arg27</param>
        /// <param name="arg28">optional object arg28</param>
        /// <param name="arg29">optional object arg29</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197597.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Select();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838617.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Show();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840851.aspx </remarks>
        /// <param name="remove">optional object remove</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object ShowDependents(object remove);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840851.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object ShowDependents();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192988.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object ShowErrors();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193004.aspx </remarks>
        /// <param name="remove">optional object remove</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object ShowPrecedents(object remove);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193004.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object ShowPrecedents();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="type">optional object type</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        /// <param name="key3">optional object key3</param>
        /// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
        /// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
        /// <param name="orderCustom">optional object orderCustom</param>
        /// <param name="matchCase">optional object matchCase</param>
        /// <param name="orientation">optional NetOffice.ExcelApi.Enums.XlSortOrientation Orientation = 2</param>
        /// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object sortMethod);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="type">optional object type</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        /// <param name="key3">optional object key3</param>
        /// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
        /// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
        /// <param name="orderCustom">optional object orderCustom</param>
        /// <param name="matchCase">optional object matchCase</param>
        /// <param name="orientation">optional NetOffice.ExcelApi.Enums.XlSortOrientation Orientation = 2</param>
        /// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
        /// <param name="dataOption1">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption1 = 0</param>
        /// <param name="dataOption2">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption2 = 0</param>
        /// <param name="dataOption3">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption3 = 0</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object sortMethod, object dataOption1, object dataOption2, object dataOption3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Sort();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
        /// <param name="key1">optional object key1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Sort(object key1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Sort(object key1, object order1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="key2">optional object key2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Sort(object key1, object order1, object key2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="type">optional object type</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Sort(object key1, object order1, object key2, object type);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="type">optional object type</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Sort(object key1, object order1, object key2, object type, object order2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="type">optional object type</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        /// <param name="key3">optional object key3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Sort(object key1, object order1, object key2, object type, object order2, object key3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="type">optional object type</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        /// <param name="key3">optional object key3</param>
        /// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="type">optional object type</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        /// <param name="key3">optional object key3</param>
        /// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
        /// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="type">optional object type</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        /// <param name="key3">optional object key3</param>
        /// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
        /// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
        /// <param name="orderCustom">optional object orderCustom</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="type">optional object type</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        /// <param name="key3">optional object key3</param>
        /// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
        /// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
        /// <param name="orderCustom">optional object orderCustom</param>
        /// <param name="matchCase">optional object matchCase</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom, object matchCase);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="type">optional object type</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        /// <param name="key3">optional object key3</param>
        /// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
        /// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
        /// <param name="orderCustom">optional object orderCustom</param>
        /// <param name="matchCase">optional object matchCase</param>
        /// <param name="orientation">optional NetOffice.ExcelApi.Enums.XlSortOrientation Orientation = 2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="type">optional object type</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        /// <param name="key3">optional object key3</param>
        /// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
        /// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
        /// <param name="orderCustom">optional object orderCustom</param>
        /// <param name="matchCase">optional object matchCase</param>
        /// <param name="orientation">optional NetOffice.ExcelApi.Enums.XlSortOrientation Orientation = 2</param>
        /// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
        /// <param name="dataOption1">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption1 = 0</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object sortMethod, object dataOption1);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="type">optional object type</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        /// <param name="key3">optional object key3</param>
        /// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
        /// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
        /// <param name="orderCustom">optional object orderCustom</param>
        /// <param name="matchCase">optional object matchCase</param>
        /// <param name="orientation">optional NetOffice.ExcelApi.Enums.XlSortOrientation Orientation = 2</param>
        /// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
        /// <param name="dataOption1">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption1 = 0</param>
        /// <param name="dataOption2">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption2 = 0</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object sortMethod, object dataOption1, object dataOption2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
        /// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="type">optional object type</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        /// <param name="key3">optional object key3</param>
        /// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
        /// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
        /// <param name="orderCustom">optional object orderCustom</param>
        /// <param name="matchCase">optional object matchCase</param>
        /// <param name="orientation">optional NetOffice.ExcelApi.Enums.XlSortOrientation Orientation = 2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
        /// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="type">optional object type</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        /// <param name="key3">optional object key3</param>
        /// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
        /// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
        /// <param name="orderCustom">optional object orderCustom</param>
        /// <param name="matchCase">optional object matchCase</param>
        /// <param name="orientation">optional NetOffice.ExcelApi.Enums.XlSortOrientation Orientation = 2</param>
        /// <param name="dataOption1">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption1 = 0</param>
        /// <param name="dataOption2">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption2 = 0</param>
        /// <param name="dataOption3">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption3 = 0</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object dataOption1, object dataOption2, object dataOption3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object SortSpecial();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
        /// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object SortSpecial(object sortMethod);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
        /// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
        /// <param name="key1">optional object key1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object SortSpecial(object sortMethod, object key1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
        /// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object SortSpecial(object sortMethod, object key1, object order1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
        /// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="type">optional object type</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object SortSpecial(object sortMethod, object key1, object order1, object type);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
        /// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="type">optional object type</param>
        /// <param name="key2">optional object key2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object SortSpecial(object sortMethod, object key1, object order1, object type, object key2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
        /// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="type">optional object type</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
        /// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="type">optional object type</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        /// <param name="key3">optional object key3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
        /// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="type">optional object type</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        /// <param name="key3">optional object key3</param>
        /// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
        /// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="type">optional object type</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        /// <param name="key3">optional object key3</param>
        /// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
        /// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
        /// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="type">optional object type</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        /// <param name="key3">optional object key3</param>
        /// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
        /// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
        /// <param name="orderCustom">optional object orderCustom</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header, object orderCustom);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
        /// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="type">optional object type</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        /// <param name="key3">optional object key3</param>
        /// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
        /// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
        /// <param name="orderCustom">optional object orderCustom</param>
        /// <param name="matchCase">optional object matchCase</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header, object orderCustom, object matchCase);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
        /// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="type">optional object type</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        /// <param name="key3">optional object key3</param>
        /// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
        /// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
        /// <param name="orderCustom">optional object orderCustom</param>
        /// <param name="matchCase">optional object matchCase</param>
        /// <param name="orientation">optional NetOffice.ExcelApi.Enums.XlSortOrientation Orientation = 2</param>
        /// <param name="dataOption1">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption1 = 0</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object dataOption1);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
        /// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="type">optional object type</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        /// <param name="key3">optional object key3</param>
        /// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
        /// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
        /// <param name="orderCustom">optional object orderCustom</param>
        /// <param name="matchCase">optional object matchCase</param>
        /// <param name="orientation">optional NetOffice.ExcelApi.Enums.XlSortOrientation Orientation = 2</param>
        /// <param name="dataOption1">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption1 = 0</param>
        /// <param name="dataOption2">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption2 = 0</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object dataOption1, object dataOption2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196157.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlCellType type</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range SpecialCells(NetOffice.ExcelApi.Enums.XlCellType type, object value);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196157.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlCellType type</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range SpecialCells(NetOffice.ExcelApi.Enums.XlCellType type);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821663.aspx </remarks>
        /// <param name="edition">string edition</param>
        /// <param name="format">optional NetOffice.ExcelApi.Enums.XlSubscribeToFormat Format = -4158</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object SubscribeTo(string edition, object format);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821663.aspx </remarks>
        /// <param name="edition">string edition</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object SubscribeTo(string edition);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838166.aspx </remarks>
        /// <param name="groupBy">Int32 groupBy</param>
        /// <param name="function">NetOffice.ExcelApi.Enums.XlConsolidationFunction function</param>
        /// <param name="totalList">object totalList</param>
        /// <param name="replace">optional object replace</param>
        /// <param name="pageBreaks">optional object pageBreaks</param>
        /// <param name="summaryBelowData">optional NetOffice.ExcelApi.Enums.XlSummaryRow SummaryBelowData = 1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Subtotal(Int32 groupBy, NetOffice.ExcelApi.Enums.XlConsolidationFunction function, object totalList, object replace, object pageBreaks, object summaryBelowData);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838166.aspx </remarks>
        /// <param name="groupBy">Int32 groupBy</param>
        /// <param name="function">NetOffice.ExcelApi.Enums.XlConsolidationFunction function</param>
        /// <param name="totalList">object totalList</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Subtotal(Int32 groupBy, NetOffice.ExcelApi.Enums.XlConsolidationFunction function, object totalList);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838166.aspx </remarks>
        /// <param name="groupBy">Int32 groupBy</param>
        /// <param name="function">NetOffice.ExcelApi.Enums.XlConsolidationFunction function</param>
        /// <param name="totalList">object totalList</param>
        /// <param name="replace">optional object replace</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Subtotal(Int32 groupBy, NetOffice.ExcelApi.Enums.XlConsolidationFunction function, object totalList, object replace);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838166.aspx </remarks>
        /// <param name="groupBy">Int32 groupBy</param>
        /// <param name="function">NetOffice.ExcelApi.Enums.XlConsolidationFunction function</param>
        /// <param name="totalList">object totalList</param>
        /// <param name="replace">optional object replace</param>
        /// <param name="pageBreaks">optional object pageBreaks</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Subtotal(Int32 groupBy, NetOffice.ExcelApi.Enums.XlConsolidationFunction function, object totalList, object replace, object pageBreaks);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834729.aspx </remarks>
        /// <param name="rowInput">optional object rowInput</param>
        /// <param name="columnInput">optional object columnInput</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Table(object rowInput, object columnInput);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834729.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Table();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834729.aspx </remarks>
        /// <param name="rowInput">optional object rowInput</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Table(object rowInput);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
        /// <param name="destination">optional object destination</param>
        /// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
        /// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
        /// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
        /// <param name="tab">optional object tab</param>
        /// <param name="semicolon">optional object semicolon</param>
        /// <param name="comma">optional object comma</param>
        /// <param name="space">optional object space</param>
        /// <param name="other">optional object other</param>
        /// <param name="otherChar">optional object otherChar</param>
        /// <param name="fieldInfo">optional object fieldInfo</param>
        /// <param name="decimalSeparator">optional object decimalSeparator</param>
        /// <param name="thousandsSeparator">optional object thousandsSeparator</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object decimalSeparator, object thousandsSeparator);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
        /// <param name="destination">optional object destination</param>
        /// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
        /// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
        /// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
        /// <param name="tab">optional object tab</param>
        /// <param name="semicolon">optional object semicolon</param>
        /// <param name="comma">optional object comma</param>
        /// <param name="space">optional object space</param>
        /// <param name="other">optional object other</param>
        /// <param name="otherChar">optional object otherChar</param>
        /// <param name="fieldInfo">optional object fieldInfo</param>
        /// <param name="decimalSeparator">optional object decimalSeparator</param>
        /// <param name="thousandsSeparator">optional object thousandsSeparator</param>
        /// <param name="trailingMinusNumbers">optional object trailingMinusNumbers</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object decimalSeparator, object thousandsSeparator, object trailingMinusNumbers);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object TextToColumns();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
        /// <param name="destination">optional object destination</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object TextToColumns(object destination);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
        /// <param name="destination">optional object destination</param>
        /// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object TextToColumns(object destination, object dataType);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
        /// <param name="destination">optional object destination</param>
        /// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
        /// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object TextToColumns(object destination, object dataType, object textQualifier);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
        /// <param name="destination">optional object destination</param>
        /// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
        /// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
        /// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
        /// <param name="destination">optional object destination</param>
        /// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
        /// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
        /// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
        /// <param name="tab">optional object tab</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
        /// <param name="destination">optional object destination</param>
        /// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
        /// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
        /// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
        /// <param name="tab">optional object tab</param>
        /// <param name="semicolon">optional object semicolon</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
        /// <param name="destination">optional object destination</param>
        /// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
        /// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
        /// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
        /// <param name="tab">optional object tab</param>
        /// <param name="semicolon">optional object semicolon</param>
        /// <param name="comma">optional object comma</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
        /// <param name="destination">optional object destination</param>
        /// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
        /// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
        /// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
        /// <param name="tab">optional object tab</param>
        /// <param name="semicolon">optional object semicolon</param>
        /// <param name="comma">optional object comma</param>
        /// <param name="space">optional object space</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
        /// <param name="destination">optional object destination</param>
        /// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
        /// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
        /// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
        /// <param name="tab">optional object tab</param>
        /// <param name="semicolon">optional object semicolon</param>
        /// <param name="comma">optional object comma</param>
        /// <param name="space">optional object space</param>
        /// <param name="other">optional object other</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
        /// <param name="destination">optional object destination</param>
        /// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
        /// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
        /// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
        /// <param name="tab">optional object tab</param>
        /// <param name="semicolon">optional object semicolon</param>
        /// <param name="comma">optional object comma</param>
        /// <param name="space">optional object space</param>
        /// <param name="other">optional object other</param>
        /// <param name="otherChar">optional object otherChar</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
        /// <param name="destination">optional object destination</param>
        /// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
        /// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
        /// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
        /// <param name="tab">optional object tab</param>
        /// <param name="semicolon">optional object semicolon</param>
        /// <param name="comma">optional object comma</param>
        /// <param name="space">optional object space</param>
        /// <param name="other">optional object other</param>
        /// <param name="otherChar">optional object otherChar</param>
        /// <param name="fieldInfo">optional object fieldInfo</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
        /// <param name="destination">optional object destination</param>
        /// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
        /// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
        /// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
        /// <param name="tab">optional object tab</param>
        /// <param name="semicolon">optional object semicolon</param>
        /// <param name="comma">optional object comma</param>
        /// <param name="space">optional object space</param>
        /// <param name="other">optional object other</param>
        /// <param name="otherChar">optional object otherChar</param>
        /// <param name="fieldInfo">optional object fieldInfo</param>
        /// <param name="decimalSeparator">optional object decimalSeparator</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object decimalSeparator);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837754.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Ungroup();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835301.aspx </remarks>
        /// <param name="text">optional object text</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Comment AddComment(object text);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835301.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Comment AddComment();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823037.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void ClearComments();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822349.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void SetPhonetic();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197435.aspx </remarks>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="prToFileName">optional object prToFileName</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197435.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object PrintOut();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197435.aspx </remarks>
        /// <param name="from">optional object from</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object PrintOut(object from);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197435.aspx </remarks>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object PrintOut(object from, object to);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197435.aspx </remarks>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object PrintOut(object from, object to, object copies);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197435.aspx </remarks>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object PrintOut(object from, object to, object copies, object preview);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197435.aspx </remarks>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object PrintOut(object from, object to, object copies, object preview, object activePrinter);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197435.aspx </remarks>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197435.aspx </remarks>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
        /// <param name="operation">optional NetOffice.ExcelApi.Enums.XlPasteSpecialOperation Operation = -4142</param>
        /// <param name="skipBlanks">optional object skipBlanks</param>
        /// <param name="transpose">optional object transpose</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        object _PasteSpecial(object paste, object operation, object skipBlanks, object transpose);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        object _PasteSpecial();

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        object _PasteSpecial(object paste);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
        /// <param name="operation">optional NetOffice.ExcelApi.Enums.XlPasteSpecialOperation Operation = -4142</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        object _PasteSpecial(object paste, object operation);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
        /// <param name="operation">optional NetOffice.ExcelApi.Enums.XlPasteSpecialOperation Operation = -4142</param>
        /// <param name="skipBlanks">optional object skipBlanks</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        object _PasteSpecial(object paste, object operation, object skipBlanks);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838797.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void Dirty();

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194108.aspx </remarks>
        /// <param name="speakDirection">optional object speakDirection</param>
        /// <param name="speakFormulas">optional object speakFormulas</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void Speak(object speakDirection, object speakFormulas);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194108.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void Speak();

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194108.aspx </remarks>
        /// <param name="speakDirection">optional object speakDirection</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void Speak(object speakDirection);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        object __PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        object __PrintOut();

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        object __PrintOut(object from);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        object __PrintOut(object from, object to);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        object __PrintOut(object from, object to, object copies);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        object __PrintOut(object from, object to, object copies, object preview);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        object __PrintOut(object from, object to, object copies, object preview, object activePrinter);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        object __PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193823.aspx </remarks>
        /// <param name="columns">optional object columns</param>
        /// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void RemoveDuplicates(object columns, object header);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193823.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void RemoveDuplicates();

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193823.aspx </remarks>
        /// <param name="columns">optional object columns</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void RemoveDuplicates(object columns);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836441.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        /// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="openAfterPublish">optional object openAfterPublish</param>
        /// <param name="fixedFormatExtClassPtr">optional object fixedFormatExtClassPtr</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to, object openAfterPublish, object fixedFormatExtClassPtr);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836441.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836441.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836441.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836441.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836441.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        /// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836441.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        /// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
        /// <param name="from">optional object from</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836441.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        /// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836441.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        /// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="openAfterPublish">optional object openAfterPublish</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to, object openAfterPublish);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835226.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        object CalculateRowMajorOrder();

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="lineStyle">optional object lineStyle</param>
        /// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
        /// <param name="colorIndex">optional NetOffice.ExcelApi.Enums.XlColorIndex ColorIndex = -4105</param>
        /// <param name="color">optional object color</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 14, 15, 16)]
        object _BorderAround(object lineStyle, object weight, object colorIndex, object color);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        object _BorderAround();

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="lineStyle">optional object lineStyle</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        object _BorderAround(object lineStyle);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="lineStyle">optional object lineStyle</param>
        /// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        object _BorderAround(object lineStyle, object weight);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="lineStyle">optional object lineStyle</param>
        /// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
        /// <param name="colorIndex">optional NetOffice.ExcelApi.Enums.XlColorIndex ColorIndex = -4105</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        object _BorderAround(object lineStyle, object weight, object colorIndex);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194741.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        void ClearHyperlinks();

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838963.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        void AllocateChanges();

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837815.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        void DiscardChanges();

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228715.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        void FlashFill();

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.Range>

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        new IEnumerator<NetOffice.ExcelApi.Range> GetEnumerator();

        #endregion
    }
}
