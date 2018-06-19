using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
    /// <summary>
    /// Table
    /// </summary>
    [SyntaxBypass]
    public interface Table_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="startRow">optional Int32 startRow</param>
        /// <param name="startColumn">optional Int32 startColumn</param>
        /// <param name="endRow">optional Int32 endRow</param>
        /// <param name="endColumn">optional Int32 endColumn</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.PublisherApi.CellRange get_Cells(object startRow, object startColumn, object endRow, object endColumn);

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Alias for get_Cells
        /// </summary>
        /// <param name="startRow">optional Int32 startRow</param>
        /// <param name="startColumn">optional Int32 startColumn</param>
        /// <param name="endRow">optional Int32 endRow</param>
        /// <param name="endColumn">optional Int32 endColumn</param>
        [SupportByVersion("Publisher", 14, 15, 16), Redirect("get_Cells")]
        NetOffice.PublisherApi.CellRange Cells(object startRow, object startColumn, object endRow, object endColumn);

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="startRow">optional Int32 startRow</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.PublisherApi.CellRange get_Cells(object startRow);

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Alias for get_Cells
        /// </summary>
        /// <param name="startRow">optional Int32 startRow</param>
        [SupportByVersion("Publisher", 14, 15, 16), Redirect("get_Cells")]
        NetOffice.PublisherApi.CellRange Cells(object startRow);

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="startRow">optional Int32 startRow</param>
        /// <param name="startColumn">optional Int32 startColumn</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.PublisherApi.CellRange get_Cells(object startRow, object startColumn);

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Alias for get_Cells
        /// </summary>
        /// <param name="startRow">optional Int32 startRow</param>
        /// <param name="startColumn">optional Int32 startColumn</param>
        [SupportByVersion("Publisher", 14, 15, 16), Redirect("get_Cells")]
        NetOffice.PublisherApi.CellRange Cells(object startRow, object startColumn);

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="startRow">optional Int32 startRow</param>
        /// <param name="startColumn">optional Int32 startColumn</param>
        /// <param name="endRow">optional Int32 endRow</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.PublisherApi.CellRange get_Cells(object startRow, object startColumn, object endRow);

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Alias for get_Cells
        /// </summary>
        /// <param name="startRow">optional Int32 startRow</param>
        /// <param name="startColumn">optional Int32 startColumn</param>
        /// <param name="endRow">optional Int32 endRow</param>
        [SupportByVersion("Publisher", 14, 15, 16), Redirect("get_Cells")]
        NetOffice.PublisherApi.CellRange Cells(object startRow, object startColumn, object endRow);

        #endregion
    }
 
    /// <summary>
    /// DispatchInterface Table 
    /// SupportByVersion Publisher, 14,15,16
    /// </summary>
    [SupportByVersion("Publisher", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("37FAE3EA-D9B4-11D3-907A-00C04F799E3F")]
    public interface Table : Table_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        NetOffice.PublisherApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        NetOffice.PublisherApi.Columns Columns { get; }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        bool GrowToFitText { get; set; }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        NetOffice.PublisherApi.Rows Rows { get; }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        NetOffice.PublisherApi.Enums.PbTableDirectionType TableDirection { get; set; }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        new NetOffice.PublisherApi.CellRange Cells { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="autoFormat">NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat</param>
        /// <param name="textFormatting">optional bool TextFormatting = true</param>
        /// <param name="textAlignment">optional bool TextAlignment = true</param>
        /// <param name="fill">optional bool Fill = true</param>
        /// <param name="borders">optional bool Borders = true</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        void ApplyAutoFormat(NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat, object textFormatting, object textAlignment, object fill, object borders);

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="autoFormat">NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat</param>
        [CustomMethod]
        [SupportByVersion("Publisher", 14, 15, 16)]
        void ApplyAutoFormat(NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat);

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="autoFormat">NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat</param>
        /// <param name="textFormatting">optional bool TextFormatting = true</param>
        [CustomMethod]
        [SupportByVersion("Publisher", 14, 15, 16)]
        void ApplyAutoFormat(NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat, object textFormatting);

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="autoFormat">NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat</param>
        /// <param name="textFormatting">optional bool TextFormatting = true</param>
        /// <param name="textAlignment">optional bool TextAlignment = true</param>
        [CustomMethod]
        [SupportByVersion("Publisher", 14, 15, 16)]
        void ApplyAutoFormat(NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat, object textFormatting, object textAlignment);

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="autoFormat">NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat</param>
        /// <param name="textFormatting">optional bool TextFormatting = true</param>
        /// <param name="textAlignment">optional bool TextAlignment = true</param>
        /// <param name="fill">optional bool Fill = true</param>
        [CustomMethod]
        [SupportByVersion("Publisher", 14, 15, 16)]
        void ApplyAutoFormat(NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat, object textFormatting, object textAlignment, object fill);

        #endregion
    }
}
