using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi.Behind
{
    /// <summary>
    /// Table
    /// </summary>
    [SyntaxBypass]
    public class Table_ : COMObject, NetOffice.PublisherApi.Table_
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public Table_() : base()
        {
        }

        #endregion

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
        public NetOffice.PublisherApi.CellRange get_Cells(object startRow, object startColumn, object endRow, object endColumn)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.CellRange>(this, "Cells", typeof(NetOffice.PublisherApi.CellRange), startRow, startColumn, endRow, endColumn);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Alias for get_Cells
        /// </summary>
        /// <param name="startRow">optional Int32 startRow</param>
        /// <param name="startColumn">optional Int32 startColumn</param>
        /// <param name="endRow">optional Int32 endRow</param>
        /// <param name="endColumn">optional Int32 endColumn</param>
        [SupportByVersion("Publisher", 14, 15, 16), Redirect("get_Cells")]
        public NetOffice.PublisherApi.CellRange Cells(object startRow, object startColumn, object endRow, object endColumn)
        {
            return get_Cells(startRow, startColumn, endRow, endColumn);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="startRow">optional Int32 startRow</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public NetOffice.PublisherApi.CellRange get_Cells(object startRow)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.CellRange>(this, "Cells", typeof(NetOffice.PublisherApi.CellRange), startRow);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Alias for get_Cells
        /// </summary>
        /// <param name="startRow">optional Int32 startRow</param>
        [SupportByVersion("Publisher", 14, 15, 16), Redirect("get_Cells")]
        public NetOffice.PublisherApi.CellRange Cells(object startRow)
        {
            return get_Cells(startRow);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="startRow">optional Int32 startRow</param>
        /// <param name="startColumn">optional Int32 startColumn</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public NetOffice.PublisherApi.CellRange get_Cells(object startRow, object startColumn)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.CellRange>(this, "Cells", typeof(NetOffice.PublisherApi.CellRange), startRow, startColumn);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Alias for get_Cells
        /// </summary>
        /// <param name="startRow">optional Int32 startRow</param>
        /// <param name="startColumn">optional Int32 startColumn</param>
        [SupportByVersion("Publisher", 14, 15, 16), Redirect("get_Cells")]
        public NetOffice.PublisherApi.CellRange Cells(object startRow, object startColumn)
        {
            return get_Cells(startRow, startColumn);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="startRow">optional Int32 startRow</param>
        /// <param name="startColumn">optional Int32 startColumn</param>
        /// <param name="endRow">optional Int32 endRow</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public NetOffice.PublisherApi.CellRange get_Cells(object startRow, object startColumn, object endRow)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.CellRange>(this, "Cells", typeof(NetOffice.PublisherApi.CellRange), startRow, startColumn, endRow);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Alias for get_Cells
        /// </summary>
        /// <param name="startRow">optional Int32 startRow</param>
        /// <param name="startColumn">optional Int32 startColumn</param>
        /// <param name="endRow">optional Int32 endRow</param>
        [SupportByVersion("Publisher", 14, 15, 16), Redirect("get_Cells")]
        public NetOffice.PublisherApi.CellRange Cells(object startRow, object startColumn, object endRow)
        {
            return get_Cells(startRow, startColumn, endRow);
        }

        #endregion
    }

    /// <summary>
    /// DispatchInterface Table 
    /// SupportByVersion Publisher, 14,15,16
    /// </summary>
    [SupportByVersion("Publisher", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class Table : Table_, NetOffice.PublisherApi.Table
    {
        #pragma warning disable

        #region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.PublisherApi.Table);
                return _contractType;
            }
        }
        private static Type _contractType;


        /// <summary>
        /// Instance Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type InstanceType
        {
            get
            {
                return LateBindingApiWrapperType;
            }
        }

        private static Type _type;

        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(Table);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public Table() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public NetOffice.PublisherApi.Application Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Application>(this, "Application", typeof(NetOffice.PublisherApi.Application));
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public NetOffice.PublisherApi.Columns Columns
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Columns>(this, "Columns", typeof(NetOffice.PublisherApi.Columns));
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public bool GrowToFitText
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "GrowToFitText");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GrowToFitText", value);
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16), ProxyResult]
        public object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public NetOffice.PublisherApi.Rows Rows
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Rows>(this, "Rows", typeof(NetOffice.PublisherApi.Rows));
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public NetOffice.PublisherApi.Enums.PbTableDirectionType TableDirection
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbTableDirectionType>(this, "TableDirection");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "TableDirection", value);
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public NetOffice.PublisherApi.CellRange Cells
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.CellRange>(this, "Cells", typeof(NetOffice.PublisherApi.CellRange));
            }
        }

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
        public void ApplyAutoFormat(NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat, object textFormatting, object textAlignment, object fill, object borders)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyAutoFormat", new object[] { autoFormat, textFormatting, textAlignment, fill, borders });
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="autoFormat">NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat</param>
        [CustomMethod]
        [SupportByVersion("Publisher", 14, 15, 16)]
        public void ApplyAutoFormat(NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyAutoFormat", autoFormat);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="autoFormat">NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat</param>
        /// <param name="textFormatting">optional bool TextFormatting = true</param>
        [CustomMethod]
        [SupportByVersion("Publisher", 14, 15, 16)]
        public void ApplyAutoFormat(NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat, object textFormatting)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyAutoFormat", autoFormat, textFormatting);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="autoFormat">NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat</param>
        /// <param name="textFormatting">optional bool TextFormatting = true</param>
        /// <param name="textAlignment">optional bool TextAlignment = true</param>
        [CustomMethod]
        [SupportByVersion("Publisher", 14, 15, 16)]
        public void ApplyAutoFormat(NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat, object textFormatting, object textAlignment)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyAutoFormat", autoFormat, textFormatting, textAlignment);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="autoFormat">NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat</param>
        /// <param name="textFormatting">optional bool TextFormatting = true</param>
        /// <param name="textAlignment">optional bool TextAlignment = true</param>
        /// <param name="fill">optional bool Fill = true</param>
        [CustomMethod]
        [SupportByVersion("Publisher", 14, 15, 16)]
        public void ApplyAutoFormat(NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat, object textFormatting, object textAlignment, object fill)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyAutoFormat", autoFormat, textFormatting, textAlignment, fill);
        }

        #endregion

        #pragma warning restore
    }
}
