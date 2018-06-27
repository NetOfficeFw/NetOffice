using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi.Behind
{
    /// <summary>
    /// LPVISIOSHAPE
    /// </summary>
    [SyntaxBypass]
    public class LPVISIOSHAPE_ : COMObject, NetOffice.VisioApi.LPVISIOSHAPE_
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public LPVISIOSHAPE_() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="fIncludeSubShapes">optional bool fIncludeSubShapes</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Double get_AreaIU(object fIncludeSubShapes)
        {
            return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "AreaIU", fIncludeSubShapes);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_AreaIU
        /// </summary>
        /// <param name="fIncludeSubShapes">optional bool fIncludeSubShapes</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_AreaIU")]
        public Double AreaIU(object fIncludeSubShapes)
        {
            return get_AreaIU(fIncludeSubShapes);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="fIncludeSubShapes">optional bool fIncludeSubShapes</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Double get_LengthIU(object fIncludeSubShapes)
        {
            return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "LengthIU", fIncludeSubShapes);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_LengthIU
        /// </summary>
        /// <param name="fIncludeSubShapes">optional bool fIncludeSubShapes</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_LengthIU")]
        public Double LengthIU(object fIncludeSubShapes)
        {
            return get_LengthIU(fIncludeSubShapes);
        }

        #endregion
    }

    /// <summary>
    /// Interface LPVISIOSHAPE 
    /// SupportByVersion Visio, 11,12,14,15,16
    /// </summary>
    [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class LPVISIOSHAPE : LPVISIOSHAPE_, NetOffice.VisioApi.LPVISIOSHAPE
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
                    _contractType = typeof(NetOffice.VisioApi.LPVISIOSHAPE);
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
                    _type = typeof(LPVISIOSHAPE);
                return _type;
            }
        }

        #endregion

        #region Ctor
 
        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public LPVISIOSHAPE() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVDocument Document
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDocument>(this, "Document");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), ProxyResult]
        public object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Parent", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVApplication Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 Stat
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Stat");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVMaster Master
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVMaster>(this, "Master");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 Type
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Type");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 ObjectType
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ObjectType");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="localeSpecificCellName">string localeSpecificCellName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public NetOffice.VisioApi.IVCell get_Cells(string localeSpecificCellName)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVCell>(this, "Cells", typeof(NetOffice.VisioApi.IVCell), localeSpecificCellName);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_Cells
        /// </summary>
        /// <param name="localeSpecificCellName">string localeSpecificCellName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_Cells")]
        public NetOffice.VisioApi.IVCell Cells(string localeSpecificCellName)
        {
            return get_Cells(localeSpecificCellName);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        /// <param name="column">Int16 column</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public NetOffice.VisioApi.IVCell get_CellsSRC(Int16 section, Int16 row, Int16 column)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVCell>(this, "CellsSRC", typeof(NetOffice.VisioApi.IVCell), section, row, column);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_CellsSRC
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        /// <param name="column">Int16 column</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_CellsSRC")]
        public NetOffice.VisioApi.IVCell CellsSRC(Int16 section, Int16 row, Int16 column)
        {
            return get_CellsSRC(section, row, column);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVShapes Shapes
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShapes>(this, "Shapes");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string Data1
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Data1");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Data1", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string Data2
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Data2");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Data2", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string Data3
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Data3");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Data3", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string Help
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Help");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Help", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string NameID
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "NameID");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string Name
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Name", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string Text
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int32 CharCount
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CharCount");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVCharacters Characters
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVCharacters>(this, "Characters");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 OneD
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "OneD");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OneD", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 GeometryCount
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "GeometryCount");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="section">Int16 section</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Int16 get_RowCount(Int16 section)
        {
            return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "RowCount", section);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_RowCount
        /// </summary>
        /// <param name="section">Int16 section</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_RowCount")]
        public Int16 RowCount(Int16 section)
        {
            return get_RowCount(section);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Int16 get_RowsCellCount(Int16 section, Int16 row)
        {
            return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "RowsCellCount", section, row);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_RowsCellCount
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_RowsCellCount")]
        public Int16 RowsCellCount(Int16 section, Int16 row)
        {
            return get_RowsCellCount(section, row);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Int16 get_RowType(Int16 section, Int16 row)
        {
            return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "RowType", section, row);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void set_RowType(Int16 section, Int16 row, Int16 value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "RowType", section, row, value);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_RowType
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_RowType")]
        public Int16 RowType(Int16 section, Int16 row)
        {
            return get_RowType(section, row);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVConnects Connects
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVConnects>(this, "Connects");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Int16 Index16
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Index16");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string Style
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Style");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Style", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string StyleKeepFmt
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "StyleKeepFmt");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "StyleKeepFmt", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string LineStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "LineStyle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LineStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string LineStyleKeepFmt
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "LineStyleKeepFmt");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LineStyleKeepFmt", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string FillStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FillStyle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FillStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string FillStyleKeepFmt
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FillStyleKeepFmt");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FillStyleKeepFmt", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string TextStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "TextStyle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TextStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string TextStyleKeepFmt
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "TextStyleKeepFmt");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TextStyleKeepFmt", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Double old_AreaIU
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "old_AreaIU");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Double old_LengthIU
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "old_LengthIU");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="fFill">Int16 fFill</param>
        /// <param name="lineRes">Double lineRes</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public object get_GeomExIf(Int16 fFill, Double lineRes)
        {
            return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "GeomExIf", fFill, lineRes);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_GeomExIf
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="fFill">Int16 fFill</param>
        /// <param name="lineRes">Double lineRes</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), ProxyResult, Redirect("get_GeomExIf")]
        public object GeomExIf(Int16 fFill, Double lineRes)
        {
            return get_GeomExIf(fFill, lineRes);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="fUniqueID">Int16 fUniqueID</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public string get_UniqueID(Int16 fUniqueID)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "UniqueID", fUniqueID);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_UniqueID
        /// </summary>
        /// <param name="fUniqueID">Int16 fUniqueID</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_UniqueID")]
        public string UniqueID(Int16 fUniqueID)
        {
            return get_UniqueID(fUniqueID);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVPage ContainingPage
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVPage>(this, "ContainingPage");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVMaster ContainingMaster
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVMaster>(this, "ContainingMaster");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVShape ContainingShape
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "ContainingShape");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="fExistsLocally">Int16 fExistsLocally</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Int16 get_SectionExists(Int16 section, Int16 fExistsLocally)
        {
            return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "SectionExists", section, fExistsLocally);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_SectionExists
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="fExistsLocally">Int16 fExistsLocally</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_SectionExists")]
        public Int16 SectionExists(Int16 section, Int16 fExistsLocally)
        {
            return get_SectionExists(section, fExistsLocally);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        /// <param name="fExistsLocally">Int16 fExistsLocally</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Int16 get_RowExists(Int16 section, Int16 row, Int16 fExistsLocally)
        {
            return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "RowExists", section, row, fExistsLocally);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_RowExists
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        /// <param name="fExistsLocally">Int16 fExistsLocally</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_RowExists")]
        public Int16 RowExists(Int16 section, Int16 row, Int16 fExistsLocally)
        {
            return get_RowExists(section, row, fExistsLocally);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="localeSpecificCellName">string localeSpecificCellName</param>
        /// <param name="fExistsLocally">Int16 fExistsLocally</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Int16 get_CellExists(string localeSpecificCellName, Int16 fExistsLocally)
        {
            return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "CellExists", localeSpecificCellName, fExistsLocally);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_CellExists
        /// </summary>
        /// <param name="localeSpecificCellName">string localeSpecificCellName</param>
        /// <param name="fExistsLocally">Int16 fExistsLocally</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_CellExists")]
        public Int16 CellExists(string localeSpecificCellName, Int16 fExistsLocally)
        {
            return get_CellExists(localeSpecificCellName, fExistsLocally);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        /// <param name="column">Int16 column</param>
        /// <param name="fExistsLocally">Int16 fExistsLocally</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Int16 get_CellsSRCExists(Int16 section, Int16 row, Int16 column, Int16 fExistsLocally)
        {
            return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "CellsSRCExists", section, row, column, fExistsLocally);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_CellsSRCExists
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        /// <param name="column">Int16 column</param>
        /// <param name="fExistsLocally">Int16 fExistsLocally</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_CellsSRCExists")]
        public Int16 CellsSRCExists(Int16 section, Int16 row, Int16 column, Int16 fExistsLocally)
        {
            return get_CellsSRCExists(section, row, column, fExistsLocally);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 LayerCount
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "LayerCount");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">Int16 index</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public NetOffice.VisioApi.IVLayer get_Layer(Int16 index)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVLayer>(this, "Layer", typeof(NetOffice.VisioApi.IVLayer), index);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_Layer
        /// </summary>
        /// <param name="index">Int16 index</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_Layer")]
        public NetOffice.VisioApi.IVLayer Layer(Int16 index)
        {
            return get_Layer(index);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVEventList EventList
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVEventList>(this, "EventList");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 PersistsEvents
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "PersistsEvents");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string ClassID
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ClassID");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 ForeignType
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ForeignType");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), ProxyResult]
        public object Object
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Object");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Int16 ID16
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ID16");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVConnects FromConnects
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVConnects>(this, "FromConnects");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public NetOffice.VisioApi.IVHyperlink Hyperlink
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVHyperlink>(this, "Hyperlink");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string ProgID
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ProgID");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 ObjectIsInherited
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ObjectIsInherited");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVPaths Paths
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVPaths>(this, "Paths");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVPaths PathsLocal
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVPaths>(this, "PathsLocal");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int32 ID
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ID");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int32 Index
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Index");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">Int16 index</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public NetOffice.VisioApi.IVSection get_Section(Int16 index)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVSection>(this, "Section", typeof(NetOffice.VisioApi.IVSection), index);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_Section
        /// </summary>
        /// <param name="index">Int16 index</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_Section")]
        public NetOffice.VisioApi.IVSection Section(Int16 index)
        {
            return get_Section(index);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVHyperlinks Hyperlinks
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVHyperlinks>(this, "Hyperlinks");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="otherShape">NetOffice.VisioApi.IVShape otherShape</param>
        /// <param name="tolerance">Double tolerance</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Int16 get_SpatialRelation(NetOffice.VisioApi.IVShape otherShape, Double tolerance, Int16 flags)
        {
            return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "SpatialRelation", otherShape, tolerance, flags);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_SpatialRelation
        /// </summary>
        /// <param name="otherShape">NetOffice.VisioApi.IVShape otherShape</param>
        /// <param name="tolerance">Double tolerance</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_SpatialRelation")]
        public Int16 SpatialRelation(NetOffice.VisioApi.IVShape otherShape, Double tolerance, Int16 flags)
        {
            return get_SpatialRelation(otherShape, tolerance, flags);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="otherShape">NetOffice.VisioApi.IVShape otherShape</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Double get_DistanceFrom(NetOffice.VisioApi.IVShape otherShape, Int16 flags)
        {
            return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "DistanceFrom", otherShape, flags);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_DistanceFrom
        /// </summary>
        /// <param name="otherShape">NetOffice.VisioApi.IVShape otherShape</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_DistanceFrom")]
        public Double DistanceFrom(NetOffice.VisioApi.IVShape otherShape, Int16 flags)
        {
            return get_DistanceFrom(otherShape, flags);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="flags">Int16 flags</param>
        /// <param name="pvPathIndex">optional object pvPathIndex</param>
        /// <param name="pvCurveIndex">optional object pvCurveIndex</param>
        /// <param name="pvt">optional object pvt</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Double get_DistanceFromPoint(Double x, Double y, Int16 flags, object pvPathIndex, object pvCurveIndex, object pvt)
        {
            return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "DistanceFromPoint", new object[] { x, y, flags, pvPathIndex, pvCurveIndex, pvt });
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_DistanceFromPoint
        /// </summary>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="flags">Int16 flags</param>
        /// <param name="pvPathIndex">optional object pvPathIndex</param>
        /// <param name="pvCurveIndex">optional object pvCurveIndex</param>
        /// <param name="pvt">optional object pvt</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_DistanceFromPoint")]
        public Double DistanceFromPoint(Double x, Double y, Int16 flags, object pvPathIndex, object pvCurveIndex, object pvt)
        {
            return get_DistanceFromPoint(x, y, flags, pvPathIndex, pvCurveIndex, pvt);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Double get_DistanceFromPoint(Double x, Double y, Int16 flags)
        {
            return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "DistanceFromPoint", x, y, flags);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_DistanceFromPoint
        /// </summary>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_DistanceFromPoint")]
        public Double DistanceFromPoint(Double x, Double y, Int16 flags)
        {
            return get_DistanceFromPoint(x, y, flags);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="flags">Int16 flags</param>
        /// <param name="pvPathIndex">optional object pvPathIndex</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Double get_DistanceFromPoint(Double x, Double y, Int16 flags, object pvPathIndex)
        {
            return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "DistanceFromPoint", x, y, flags, pvPathIndex);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_DistanceFromPoint
        /// </summary>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="flags">Int16 flags</param>
        /// <param name="pvPathIndex">optional object pvPathIndex</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_DistanceFromPoint")]
        public Double DistanceFromPoint(Double x, Double y, Int16 flags, object pvPathIndex)
        {
            return get_DistanceFromPoint(x, y, flags, pvPathIndex);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="flags">Int16 flags</param>
        /// <param name="pvPathIndex">optional object pvPathIndex</param>
        /// <param name="pvCurveIndex">optional object pvCurveIndex</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Double get_DistanceFromPoint(Double x, Double y, Int16 flags, object pvPathIndex, object pvCurveIndex)
        {
            return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "DistanceFromPoint", new object[] { x, y, flags, pvPathIndex, pvCurveIndex });
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_DistanceFromPoint
        /// </summary>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="flags">Int16 flags</param>
        /// <param name="pvPathIndex">optional object pvPathIndex</param>
        /// <param name="pvCurveIndex">optional object pvCurveIndex</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_DistanceFromPoint")]
        public Double DistanceFromPoint(Double x, Double y, Int16 flags, object pvPathIndex, object pvCurveIndex)
        {
            return get_DistanceFromPoint(x, y, flags, pvPathIndex, pvCurveIndex);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="relation">Int16 relation</param>
        /// <param name="tolerance">Double tolerance</param>
        /// <param name="flags">Int16 flags</param>
        /// <param name="resultRoot">optional object resultRoot</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public NetOffice.VisioApi.IVSelection get_SpatialNeighbors(Int16 relation, Double tolerance, Int16 flags, object resultRoot)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVSelection>(this, "SpatialNeighbors", typeof(NetOffice.VisioApi.IVSelection), relation, tolerance, flags, resultRoot);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_SpatialNeighbors
        /// </summary>
        /// <param name="relation">Int16 relation</param>
        /// <param name="tolerance">Double tolerance</param>
        /// <param name="flags">Int16 flags</param>
        /// <param name="resultRoot">optional object resultRoot</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_SpatialNeighbors")]
        public NetOffice.VisioApi.IVSelection SpatialNeighbors(Int16 relation, Double tolerance, Int16 flags, object resultRoot)
        {
            return get_SpatialNeighbors(relation, tolerance, flags, resultRoot);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="relation">Int16 relation</param>
        /// <param name="tolerance">Double tolerance</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public NetOffice.VisioApi.IVSelection get_SpatialNeighbors(Int16 relation, Double tolerance, Int16 flags)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVSelection>(this, "SpatialNeighbors", typeof(NetOffice.VisioApi.IVSelection), relation, tolerance, flags);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_SpatialNeighbors
        /// </summary>
        /// <param name="relation">Int16 relation</param>
        /// <param name="tolerance">Double tolerance</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_SpatialNeighbors")]
        public NetOffice.VisioApi.IVSelection SpatialNeighbors(Int16 relation, Double tolerance, Int16 flags)
        {
            return get_SpatialNeighbors(relation, tolerance, flags);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="relation">Int16 relation</param>
        /// <param name="tolerance">Double tolerance</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public NetOffice.VisioApi.IVSelection get_SpatialSearch(Double x, Double y, Int16 relation, Double tolerance, Int16 flags)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVSelection>(this, "SpatialSearch", typeof(NetOffice.VisioApi.IVSelection), new object[] { x, y, relation, tolerance, flags });
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_SpatialSearch
        /// </summary>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="relation">Int16 relation</param>
        /// <param name="tolerance">Double tolerance</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_SpatialSearch")]
        public NetOffice.VisioApi.IVSelection SpatialSearch(Double x, Double y, Int16 relation, Double tolerance, Int16 flags)
        {
            return get_SpatialSearch(x, y, relation, tolerance, flags);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="localeIndependentCellName">string localeIndependentCellName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public NetOffice.VisioApi.IVCell get_CellsU(string localeIndependentCellName)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVCell>(this, "CellsU", typeof(NetOffice.VisioApi.IVCell), localeIndependentCellName);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_CellsU
        /// </summary>
        /// <param name="localeIndependentCellName">string localeIndependentCellName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_CellsU")]
        public NetOffice.VisioApi.IVCell CellsU(string localeIndependentCellName)
        {
            return get_CellsU(localeIndependentCellName);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string NameU
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "NameU");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "NameU", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="localeIndependentCellName">string localeIndependentCellName</param>
        /// <param name="fExistsLocally">Int16 fExistsLocally</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Int16 get_CellExistsU(string localeIndependentCellName, Int16 fExistsLocally)
        {
            return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "CellExistsU", localeIndependentCellName, fExistsLocally);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_CellExistsU
        /// </summary>
        /// <param name="localeIndependentCellName">string localeIndependentCellName</param>
        /// <param name="fExistsLocally">Int16 fExistsLocally</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_CellExistsU")]
        public Int16 CellExistsU(string localeIndependentCellName, Int16 fExistsLocally)
        {
            return get_CellExistsU(localeIndependentCellName, fExistsLocally);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="localeSpecificCellName">string localeSpecificCellName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Int16 get_CellsRowIndex(string localeSpecificCellName)
        {
            return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "CellsRowIndex", localeSpecificCellName);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_CellsRowIndex
        /// </summary>
        /// <param name="localeSpecificCellName">string localeSpecificCellName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_CellsRowIndex")]
        public Int16 CellsRowIndex(string localeSpecificCellName)
        {
            return get_CellsRowIndex(localeSpecificCellName);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="localeIndependentCellName">string localeIndependentCellName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Int16 get_CellsRowIndexU(string localeIndependentCellName)
        {
            return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "CellsRowIndexU", localeIndependentCellName);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_CellsRowIndexU
        /// </summary>
        /// <param name="localeIndependentCellName">string localeIndependentCellName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_CellsRowIndexU")]
        public Int16 CellsRowIndexU(string localeIndependentCellName)
        {
            return get_CellsRowIndexU(localeIndependentCellName);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public bool IsOpenForTextEdit
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsOpenForTextEdit");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVShape RootShape
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "RootShape");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVShape MasterShape
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "MasterShape");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), NativeResult]
        public stdole.Picture Picture
        {
            get
            {
                object[] paramsArray = null;
                object returnItem = Invoker.PropertyGet(this, "Picture", paramsArray);
                return returnItem as stdole.Picture;
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public byte[] ForeignData
        {
            get
            {
                object[] paramsArray = null;
                object returnItem = (object)Invoker.PropertyGet(this, "ForeignData", paramsArray);
                return (byte[])returnItem;
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int32 Language
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Language");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Language", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Double AreaIU
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "AreaIU");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Double LengthIU
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "LengthIU");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int32 ContainingPageID
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ContainingPageID");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int32 ContainingMasterID
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ContainingMasterID");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVMaster DataGraphic
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVMaster>(this, "DataGraphic");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "DataGraphic", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public bool IsDataGraphicCallout
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsDataGraphicCallout");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVContainerProperties ContainerProperties
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVContainerProperties>(this, "ContainerProperties");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        public Int32[] MemberOfContainers
        {
            get
            {
                object[] paramsArray = null;
                object returnItem = (object)Invoker.PropertyGet(this, "MemberOfContainers", paramsArray);
                return (Int32[])returnItem;
            }
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        public bool IsCallout
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsCallout");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVShape CalloutTarget
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "CalloutTarget");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "CalloutTarget", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        public Int32[] CalloutsAssociated
        {
            get
            {
                object[] paramsArray = null;
                object returnItem = (object)Invoker.PropertyGet(this, "CalloutsAssociated", paramsArray);
                return (Int32[])returnItem;
            }
        }

        /// <summary>
        /// SupportByVersion Visio 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVComments Comments
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVComments>(this, "Comments");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void VoidGroup()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "VoidGroup");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void BringForward()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "BringForward");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void BringToFront()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "BringToFront");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void ConvertToGroup()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ConvertToGroup");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void FlipHorizontal()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "FlipHorizontal");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void FlipVertical()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "FlipVertical");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void ReverseEnds()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ReverseEnds");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void SendBackward()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SendBackward");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void SendToBack()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SendToBack");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Rotate90()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Rotate90");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Ungroup()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Ungroup");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void old_Copy()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "old_Copy");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void old_Cut()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "old_Cut");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Delete()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void VoidDuplicate()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "VoidDuplicate");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="objectToDrop">object objectToDrop</param>
        /// <param name="xPos">Double xPos</param>
        /// <param name="yPos">Double yPos</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVShape Drop(object objectToDrop, Double xPos, Double yPos)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "Drop", objectToDrop, xPos, yPos);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="section">Int16 section</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 AddSection(Int16 section)
        {
            return InvokerService.InvokeInternal.ExecuteInt16MethodGet(this, "AddSection", section);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="section">Int16 section</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void DeleteSection(Int16 section)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteSection", section);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        /// <param name="rowTag">Int16 rowTag</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 AddRow(Int16 section, Int16 row, Int16 rowTag)
        {
            return InvokerService.InvokeInternal.ExecuteInt16MethodGet(this, "AddRow", section, row, rowTag);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void DeleteRow(Int16 section, Int16 row)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteRow", section, row);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xPos">Double xPos</param>
        /// <param name="yPos">Double yPos</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void SetCenter(Double xPos, Double yPos)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetCenter", xPos, yPos);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xPos">Double xPos</param>
        /// <param name="yPos">Double yPos</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void SetBegin(Double xPos, Double yPos)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetBegin", xPos, yPos);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xPos">Double xPos</param>
        /// <param name="yPos">Double yPos</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void SetEnd(Double xPos, Double yPos)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetEnd", xPos, yPos);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Export(string fileName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Export", fileName);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="rowName">string rowName</param>
        /// <param name="rowTag">Int16 rowTag</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 AddNamedRow(Int16 section, string rowName, Int16 rowTag)
        {
            return InvokerService.InvokeInternal.ExecuteInt16MethodGet(this, "AddNamedRow", section, rowName, rowTag);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="section">Int16 section</param>
        /// <param name="row">Int16 row</param>
        /// <param name="rowTag">Int16 rowTag</param>
        /// <param name="rowCount">Int16 rowCount</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 AddRows(Int16 section, Int16 row, Int16 rowTag, Int16 rowCount)
        {
            return InvokerService.InvokeInternal.ExecuteInt16MethodGet(this, "AddRows", section, row, rowTag, rowCount);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xBegin">Double xBegin</param>
        /// <param name="yBegin">Double yBegin</param>
        /// <param name="xEnd">Double xEnd</param>
        /// <param name="yEnd">Double yEnd</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVShape DrawLine(Double xBegin, Double yBegin, Double xEnd, Double yEnd)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawLine", xBegin, yBegin, xEnd, yEnd);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="x1">Double x1</param>
        /// <param name="y1">Double y1</param>
        /// <param name="x2">Double x2</param>
        /// <param name="y2">Double y2</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVShape DrawRectangle(Double x1, Double y1, Double x2, Double y2)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawRectangle", x1, y1, x2, y2);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="x1">Double x1</param>
        /// <param name="y1">Double y1</param>
        /// <param name="x2">Double x2</param>
        /// <param name="y2">Double y2</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVShape DrawOval(Double x1, Double y1, Double x2, Double y2)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawOval", x1, y1, x2, y2);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xyArray">Double[] xyArray</param>
        /// <param name="tolerance">Double tolerance</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public NetOffice.VisioApi.IVShape DrawSpline(Double[] xyArray, Double tolerance, Int16 flags)
        {
            object[] paramsArray = Invoker.ValidateParamsArray((object)xyArray, tolerance, flags);
            object returnItem = Invoker.MethodReturn(this, "DrawSpline", paramsArray);
            NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy<NetOffice.VisioApi.IVShape>(this, returnItem, false);
            return newObject;
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xyArray">Double[] xyArray</param>
        /// <param name="degree">Int16 degree</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public NetOffice.VisioApi.IVShape DrawBezier(Double[] xyArray, Int16 degree, Int16 flags)
        {
            object[] paramsArray = Invoker.ValidateParamsArray((object)xyArray, degree, flags);
            object returnItem = Invoker.MethodReturn(this, "DrawBezier", paramsArray);
            NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy<NetOffice.VisioApi.IVShape>(this, returnItem, false);
            return newObject;
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xyArray">Double[] xyArray</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public NetOffice.VisioApi.IVShape DrawPolyline(Double[] xyArray, Int16 flags)
        {
            object[] paramsArray = Invoker.ValidateParamsArray((object)xyArray, flags);
            object returnItem = Invoker.MethodReturn(this, "DrawPolyline", paramsArray);
            NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy<NetOffice.VisioApi.IVShape>(this, returnItem, false);
            return newObject;
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="tolerance">Double tolerance</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void FitCurve(Double tolerance, Int16 flags)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "FitCurve", tolerance, flags);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVShape Import(string fileName)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "Import", fileName);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void CenterDrawing()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CenterDrawing");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="fileName">string fileName</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVShape InsertFromFile(string fileName, Int16 flags)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "InsertFromFile", fileName, flags);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="classOrProgID">string classOrProgID</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVShape InsertObject(string classOrProgID, Int16 flags)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "InsertObject", classOrProgID, flags);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVWindow OpenDrawWindow()
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "OpenDrawWindow");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVWindow OpenSheetWindow()
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "OpenSheetWindow");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="objectsToInstance">object[] objectsToInstance</param>
        /// <param name="xyArray">Double[] xyArray</param>
        /// <param name="iDArray">Int16[] iDArray</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 DropMany(object[] objectsToInstance, Double[] xyArray, out Int16[] iDArray)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, false, true);
            iDArray = null;
            object[] paramsArray = Invoker.ValidateParamsArray((object)objectsToInstance, (object)xyArray, (object)iDArray);
            object returnItem = Invoker.MethodReturn(this, "DropMany", paramsArray);
            iDArray = (Int16[])paramsArray[2];
            return NetRuntimeSystem.Convert.ToInt16(returnItem);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sRCStream">Int16[] sRCStream</param>
        /// <param name="formulaArray">object[] formulaArray</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void GetFormulas(Int16[] sRCStream, out object[] formulaArray)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, true);
            formulaArray = null;
            object[] paramsArray = Invoker.ValidateParamsArray((object)sRCStream, (object)formulaArray);
            Invoker.Method(this, "GetFormulas", paramsArray, modifiers);
            formulaArray = (object[])paramsArray[1];
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sRCStream">Int16[] sRCStream</param>
        /// <param name="flags">Int16 flags</param>
        /// <param name="unitsNamesOrCodes">object[] unitsNamesOrCodes</param>
        /// <param name="resultArray">object[] resultArray</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void GetResults(Int16[] sRCStream, Int16 flags, object[] unitsNamesOrCodes, out object[] resultArray)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, false, false, true);
            resultArray = null;
            object[] paramsArray = Invoker.ValidateParamsArray((object)sRCStream, flags, (object)unitsNamesOrCodes, (object)resultArray);
            Invoker.Method(this, "GetResults", paramsArray, modifiers);
            resultArray = (object[])paramsArray[3];
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sRCStream">Int16[] sRCStream</param>
        /// <param name="formulaArray">object[] formulaArray</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 SetFormulas(Int16[] sRCStream, object[] formulaArray, Int16 flags)
        {
            object[] paramsArray = Invoker.ValidateParamsArray((object)sRCStream, (object)formulaArray, flags);
            object returnItem = Invoker.MethodReturn(this, "SetFormulas", paramsArray);
            return NetRuntimeSystem.Convert.ToInt16(returnItem);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sRCStream">Int16[] sRCStream</param>
        /// <param name="unitsNamesOrCodes">object[] unitsNamesOrCodes</param>
        /// <param name="resultArray">object[] resultArray</param>
        /// <param name="flags">Int16 flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 SetResults(Int16[] sRCStream, object[] unitsNamesOrCodes, object[] resultArray, Int16 flags)
        {
            object[] paramsArray = Invoker.ValidateParamsArray((object)sRCStream, (object)unitsNamesOrCodes, (object)resultArray, flags);
            object returnItem = Invoker.MethodReturn(this, "SetResults", paramsArray);
            return NetRuntimeSystem.Convert.ToInt16(returnItem);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Layout()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Layout");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="flags">Int16 flags</param>
        /// <param name="lpr8Left">Double lpr8Left</param>
        /// <param name="lpr8Bottom">Double lpr8Bottom</param>
        /// <param name="lpr8Right">Double lpr8Right</param>
        /// <param name="lpr8Top">Double lpr8Top</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void BoundingBox(Int16 flags, out Double lpr8Left, out Double lpr8Bottom, out Double lpr8Right, out Double lpr8Top)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, true, true, true, true);
            lpr8Left = 0;
            lpr8Bottom = 0;
            lpr8Right = 0;
            lpr8Top = 0;
            object[] paramsArray = Invoker.ValidateParamsArray(flags, lpr8Left, lpr8Bottom, lpr8Right, lpr8Top);
            Invoker.Method(this, "BoundingBox", paramsArray, modifiers);
            lpr8Left = (Double)paramsArray[1];
            lpr8Bottom = (Double)paramsArray[2];
            lpr8Right = (Double)paramsArray[3];
            lpr8Top = (Double)paramsArray[4];
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xPos">Double xPos</param>
        /// <param name="yPos">Double yPos</param>
        /// <param name="tolerance">Double tolerance</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 HitTest(Double xPos, Double yPos, Double tolerance)
        {
            return InvokerService.InvokeInternal.ExecuteInt16MethodGet(this, "HitTest", xPos, yPos, tolerance);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVHyperlink AddHyperlink()
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVHyperlink>(this, "AddHyperlink");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="otherShape">NetOffice.VisioApi.IVShape otherShape</param>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="xprime">Double xprime</param>
        /// <param name="yprime">Double yprime</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void TransformXYTo(NetOffice.VisioApi.IVShape otherShape, Double x, Double y, out Double xprime, out Double yprime)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, false, false, true, true);
            xprime = 0;
            yprime = 0;
            object[] paramsArray = Invoker.ValidateParamsArray(otherShape, x, y, xprime, yprime);
            Invoker.Method(this, "TransformXYTo", paramsArray, modifiers);
            xprime = (Double)paramsArray[3];
            yprime = (Double)paramsArray[4];
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="otherShape">NetOffice.VisioApi.IVShape otherShape</param>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="xprime">Double xprime</param>
        /// <param name="yprime">Double yprime</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void TransformXYFrom(NetOffice.VisioApi.IVShape otherShape, Double x, Double y, out Double xprime, out Double yprime)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, false, false, true, true);
            xprime = 0;
            yprime = 0;
            object[] paramsArray = Invoker.ValidateParamsArray(otherShape, x, y, xprime, yprime);
            Invoker.Method(this, "TransformXYFrom", paramsArray, modifiers);
            xprime = (Double)paramsArray[3];
            yprime = (Double)paramsArray[4];
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="xprime">Double xprime</param>
        /// <param name="yprime">Double yprime</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void XYToPage(Double x, Double y, out Double xprime, out Double yprime)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, false, true, true);
            xprime = 0;
            yprime = 0;
            object[] paramsArray = Invoker.ValidateParamsArray(x, y, xprime, yprime);
            Invoker.Method(this, "XYToPage", paramsArray, modifiers);
            xprime = (Double)paramsArray[2];
            yprime = (Double)paramsArray[3];
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="x">Double x</param>
        /// <param name="y">Double y</param>
        /// <param name="xprime">Double xprime</param>
        /// <param name="yprime">Double yprime</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void XYFromPage(Double x, Double y, out Double xprime, out Double yprime)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, false, true, true);
            xprime = 0;
            yprime = 0;
            object[] paramsArray = Invoker.ValidateParamsArray(x, y, xprime, yprime);
            Invoker.Method(this, "XYFromPage", paramsArray, modifiers);
            xprime = (Double)paramsArray[2];
            yprime = (Double)paramsArray[3];
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void UpdateAlignmentBox()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateAlignmentBox");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="objectsToInstance">object[] objectsToInstance</param>
        /// <param name="xyArray">Double[] xyArray</param>
        /// <param name="iDArray">Int16[] iDArray</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 DropManyU(object[] objectsToInstance, Double[] xyArray, out Int16[] iDArray)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, false, true);
            iDArray = null;
            object[] paramsArray = Invoker.ValidateParamsArray((object)objectsToInstance, (object)xyArray, (object)iDArray);
            object returnItem = Invoker.MethodReturn(this, "DropManyU", paramsArray);
            iDArray = (Int16[])paramsArray[2];
            return NetRuntimeSystem.Convert.ToInt16(returnItem);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sRCStream">Int16[] sRCStream</param>
        /// <param name="formulaArray">object[] formulaArray</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void GetFormulasU(Int16[] sRCStream, out object[] formulaArray)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, true);
            formulaArray = null;
            object[] paramsArray = Invoker.ValidateParamsArray((object)sRCStream, (object)formulaArray);
            Invoker.Method(this, "GetFormulasU", paramsArray, modifiers);
            formulaArray = (object[])paramsArray[1];
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="degree">Int16 degree</param>
        /// <param name="flags">Int16 flags</param>
        /// <param name="xyArray">Double[] xyArray</param>
        /// <param name="knots">Double[] knots</param>
        /// <param name="weights">optional object weights</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public NetOffice.VisioApi.IVShape DrawNURBS(Int16 degree, Int16 flags, Double[] xyArray, Double[] knots, object weights)
        {
            object[] paramsArray = Invoker.ValidateParamsArray(degree, flags, (object)xyArray, (object)knots, weights);
            object returnItem = Invoker.MethodReturn(this, "DrawNURBS", paramsArray);
            NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy<NetOffice.VisioApi.IVShape>(this, returnItem, false);
            return newObject;
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="degree">Int16 degree</param>
        /// <param name="flags">Int16 flags</param>
        /// <param name="xyArray">Double[] xyArray</param>
        /// <param name="knots">Double[] knots</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public NetOffice.VisioApi.IVShape DrawNURBS(Int16 degree, Int16 flags, Double[] xyArray, Double[] knots)
        {
            object[] paramsArray = Invoker.ValidateParamsArray(degree, flags, (object)xyArray, (object)knots);
            object returnItem = Invoker.MethodReturn(this, "DrawNURBS", paramsArray);
            NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy<NetOffice.VisioApi.IVShape>(this, returnItem, false);
            return newObject;
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVShape Group()
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "Group");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVShape Duplicate()
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "Duplicate");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void SwapEnds()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SwapEnds");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="flags">optional object flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Copy(object flags)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Copy", flags);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Copy()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Copy");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="flags">optional object flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Cut(object flags)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Cut", flags);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Cut()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Cut");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="flags">optional object flags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Paste(object flags)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Paste", flags);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Paste()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Paste");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">Int32 format</param>
        /// <param name="link">optional object link</param>
        /// <param name="displayAsIcon">optional object displayAsIcon</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void PasteSpecial(Int32 format, object link, object displayAsIcon)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PasteSpecial", format, link, displayAsIcon);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">Int32 format</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void PasteSpecial(Int32 format)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PasteSpecial", format);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">Int32 format</param>
        /// <param name="link">optional object link</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void PasteSpecial(Int32 format, object link)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PasteSpecial", format, link);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="selType">NetOffice.VisioApi.Enums.VisSelectionTypes selType</param>
        /// <param name="iterationMode">optional NetOffice.VisioApi.Enums.VisSelectMode IterationMode = 256</param>
        /// <param name="data">optional object data</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVSelection CreateSelection(NetOffice.VisioApi.Enums.VisSelectionTypes selType, object iterationMode, object data)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVSelection>(this, "CreateSelection", selType, iterationMode, data);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="selType">NetOffice.VisioApi.Enums.VisSelectionTypes selType</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public NetOffice.VisioApi.IVSelection CreateSelection(NetOffice.VisioApi.Enums.VisSelectionTypes selType)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVSelection>(this, "CreateSelection", selType);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="selType">NetOffice.VisioApi.Enums.VisSelectionTypes selType</param>
        /// <param name="iterationMode">optional NetOffice.VisioApi.Enums.VisSelectMode IterationMode = 256</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public NetOffice.VisioApi.IVSelection CreateSelection(NetOffice.VisioApi.Enums.VisSelectionTypes selType, object iterationMode)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVSelection>(this, "CreateSelection", selType, iterationMode);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="distance">Double distance</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Offset(Double distance)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Offset", distance);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="type">Int16 type</param>
        /// <param name="xPos">Double xPos</param>
        /// <param name="yPos">Double yPos</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVShape AddGuide(Int16 type, Double xPos, Double yPos)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "AddGuide", type, xPos, yPos);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xBegin">Double xBegin</param>
        /// <param name="yBegin">Double yBegin</param>
        /// <param name="xEnd">Double xEnd</param>
        /// <param name="yEnd">Double yEnd</param>
        /// <param name="xControl">Double xControl</param>
        /// <param name="yControl">Double yControl</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVShape DrawArcByThreePoints(Double xBegin, Double yBegin, Double xEnd, Double yEnd, Double xControl, Double yControl)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawArcByThreePoints", new object[] { xBegin, yBegin, xEnd, yEnd, xControl, yControl });
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xBegin">Double xBegin</param>
        /// <param name="yBegin">Double yBegin</param>
        /// <param name="xEnd">Double xEnd</param>
        /// <param name="yEnd">Double yEnd</param>
        /// <param name="sweepFlag">NetOffice.VisioApi.Enums.VisArcSweepFlags sweepFlag</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVShape DrawQuarterArc(Double xBegin, Double yBegin, Double xEnd, Double yEnd, NetOffice.VisioApi.Enums.VisArcSweepFlags sweepFlag)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawQuarterArc", new object[] { xBegin, yBegin, xEnd, yEnd, sweepFlag });
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xCenter">Double xCenter</param>
        /// <param name="yCenter">Double yCenter</param>
        /// <param name="radius">Double radius</param>
        /// <param name="startAngle">optional Double StartAngle = 0</param>
        /// <param name="endAngle">optional Double EndAngle = 3.1415927410125732</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVShape DrawCircularArc(Double xCenter, Double yCenter, Double radius, object startAngle, object endAngle)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawCircularArc", new object[] { xCenter, yCenter, radius, startAngle, endAngle });
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xCenter">Double xCenter</param>
        /// <param name="yCenter">Double yCenter</param>
        /// <param name="radius">Double radius</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public NetOffice.VisioApi.IVShape DrawCircularArc(Double xCenter, Double yCenter, Double radius)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawCircularArc", xCenter, yCenter, radius);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xCenter">Double xCenter</param>
        /// <param name="yCenter">Double yCenter</param>
        /// <param name="radius">Double radius</param>
        /// <param name="startAngle">optional Double StartAngle = 0</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public NetOffice.VisioApi.IVShape DrawCircularArc(Double xCenter, Double yCenter, Double radius, object startAngle)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawCircularArc", xCenter, yCenter, radius, startAngle);
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
        /// <param name="rowID">Int32 rowID</param>
        /// <param name="applyDataGraphicAfterLink">optional bool ApplyDataGraphicAfterLink = true</param>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public void LinkToData(Int32 dataRecordsetID, Int32 rowID, object applyDataGraphicAfterLink)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "LinkToData", dataRecordsetID, rowID, applyDataGraphicAfterLink);
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
        /// <param name="rowID">Int32 rowID</param>
        [CustomMethod]
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public void LinkToData(Int32 dataRecordsetID, Int32 rowID)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "LinkToData", dataRecordsetID, rowID);
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public void BreakLinkToData(Int32 dataRecordsetID)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "BreakLinkToData", dataRecordsetID);
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public Int32 GetLinkedDataRow(Int32 dataRecordsetID)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetLinkedDataRow", dataRecordsetID);
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="dataRecordsetIDs">Int32[] dataRecordsetIDs</param>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public void GetLinkedDataRecordsetIDs(out Int32[] dataRecordsetIDs)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
            dataRecordsetIDs = null;
            object[] paramsArray = Invoker.ValidateParamsArray((object)dataRecordsetIDs);
            Invoker.Method(this, "GetLinkedDataRecordsetIDs", paramsArray, modifiers);
            dataRecordsetIDs = (Int32[])paramsArray[0];
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
        /// <param name="customPropertyIndices">Int32[] customPropertyIndices</param>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public void GetCustomPropertiesLinkedToData(Int32 dataRecordsetID, out Int32[] customPropertyIndices)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, true);
            customPropertyIndices = null;
            object[] paramsArray = Invoker.ValidateParamsArray(dataRecordsetID, (object)customPropertyIndices);
            Invoker.Method(this, "GetCustomPropertiesLinkedToData", paramsArray, modifiers);
            customPropertyIndices = (Int32[])paramsArray[1];
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
        /// <param name="customPropertyIndex">Int32 customPropertyIndex</param>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public bool IsCustomPropertyLinked(Int32 dataRecordsetID, Int32 customPropertyIndex)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsCustomPropertyLinked", dataRecordsetID, customPropertyIndex);
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
        /// <param name="customPropertyIndex">Int32 customPropertyIndex</param>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public string GetCustomPropertyLinkedColumn(Int32 dataRecordsetID, Int32 customPropertyIndex)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetCustomPropertyLinkedColumn", dataRecordsetID, customPropertyIndex);
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="toShape">NetOffice.VisioApi.IVShape toShape</param>
        /// <param name="placementDir">NetOffice.VisioApi.Enums.VisAutoConnectDir placementDir</param>
        /// <param name="connector">optional object Connector = null (Nothing in visual basic)</param>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public void AutoConnect(NetOffice.VisioApi.IVShape toShape, NetOffice.VisioApi.Enums.VisAutoConnectDir placementDir, object connector)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AutoConnect", toShape, placementDir, connector);
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="toShape">NetOffice.VisioApi.IVShape toShape</param>
        /// <param name="placementDir">NetOffice.VisioApi.Enums.VisAutoConnectDir placementDir</param>
        [CustomMethod]
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public void AutoConnect(NetOffice.VisioApi.IVShape toShape, NetOffice.VisioApi.Enums.VisAutoConnectDir placementDir)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AutoConnect", toShape, placementDir);
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        /// <param name="category">string category</param>
        [SupportByVersion("Visio", 14, 15, 16)]
        public bool HasCategory(string category)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "HasCategory", category);
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        /// <param name="flags">NetOffice.VisioApi.Enums.VisConnectedShapesFlags flags</param>
        /// <param name="categoryFilter">string categoryFilter</param>
        [SupportByVersion("Visio", 14, 15, 16)]
        public Int32[] ConnectedShapes(NetOffice.VisioApi.Enums.VisConnectedShapesFlags flags, string categoryFilter)
        {
            object[] paramsArray = Invoker.ValidateParamsArray(flags, categoryFilter);
            object returnItem = (object)Invoker.MethodReturn(this, "ConnectedShapes", paramsArray);
            return (Int32[])returnItem;
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        /// <param name="flags">NetOffice.VisioApi.Enums.VisGluedShapesFlags flags</param>
        /// <param name="categoryFilter">string categoryFilter</param>
        /// <param name="pOtherConnectedShape">optional NetOffice.VisioApi.IVShape pOtherConnectedShape</param>
        [SupportByVersion("Visio", 14, 15, 16)]
        public Int32[] GluedShapes(NetOffice.VisioApi.Enums.VisGluedShapesFlags flags, string categoryFilter, object pOtherConnectedShape)
        {
            object[] paramsArray = Invoker.ValidateParamsArray(flags, categoryFilter, pOtherConnectedShape);
            object returnItem = (object)Invoker.MethodReturn(this, "GluedShapes", paramsArray);
            return (Int32[])returnItem;
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        /// <param name="flags">NetOffice.VisioApi.Enums.VisGluedShapesFlags flags</param>
        /// <param name="categoryFilter">string categoryFilter</param>
        [CustomMethod]
        [SupportByVersion("Visio", 14, 15, 16)]
        public Int32[] GluedShapes(NetOffice.VisioApi.Enums.VisGluedShapesFlags flags, string categoryFilter)
        {
            object[] paramsArray = Invoker.ValidateParamsArray(flags, categoryFilter);
            object returnItem = (object)Invoker.MethodReturn(this, "GluedShapes", paramsArray);
            return (Int32[])returnItem;
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        /// <param name="connectorEnd">NetOffice.VisioApi.Enums.VisConnectorEnds connectorEnd</param>
        /// <param name="offsetX">Double offsetX</param>
        /// <param name="offsetY">Double offsetY</param>
        /// <param name="units">NetOffice.VisioApi.Enums.VisUnitCodes units</param>
        [SupportByVersion("Visio", 14, 15, 16)]
        public void Disconnect(NetOffice.VisioApi.Enums.VisConnectorEnds connectorEnd, Double offsetX, Double offsetY, NetOffice.VisioApi.Enums.VisUnitCodes units)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Disconnect", connectorEnd, offsetX, offsetY, units);
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        /// <param name="direction">NetOffice.VisioApi.Enums.VisResizeDirection direction</param>
        /// <param name="distance">Double distance</param>
        /// <param name="unitCode">NetOffice.VisioApi.Enums.VisUnitCodes unitCode</param>
        [SupportByVersion("Visio", 14, 15, 16)]
        public void Resize(NetOffice.VisioApi.Enums.VisResizeDirection direction, Double distance, NetOffice.VisioApi.Enums.VisUnitCodes unitCode)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Resize", direction, distance, unitCode);
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        public void AddToContainers()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AddToContainers");
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        public void RemoveFromContainers()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "RemoveFromContainers");
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVPage CreateSubProcess()
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVPage>(this, "CreateSubProcess");
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        /// <param name="page">NetOffice.VisioApi.IVPage page</param>
        /// <param name="objectToDrop">object objectToDrop</param>
        /// <param name="newShape">optional NetOffice.VisioApi.IVShape NewShape = 0</param>
        [SupportByVersion("Visio", 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVSelection MoveToSubprocess(NetOffice.VisioApi.IVPage page, object objectToDrop, object newShape)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVSelection>(this, "MoveToSubprocess", page, objectToDrop, newShape);
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        /// <param name="page">NetOffice.VisioApi.IVPage page</param>
        /// <param name="objectToDrop">object objectToDrop</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("Visio", 14, 15, 16)]
        public NetOffice.VisioApi.IVSelection MoveToSubprocess(NetOffice.VisioApi.IVPage page, object objectToDrop)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVSelection>(this, "MoveToSubprocess", page, objectToDrop);
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// </summary>
        /// <param name="delFlags">Int32 delFlags</param>
        [SupportByVersion("Visio", 14, 15, 16)]
        public void DeleteEx(Int32 delFlags)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteEx", delFlags);
        }

        /// <summary>
        /// SupportByVersion Visio 15,16
        /// </summary>
        /// <param name="masterOrMasterShortcutToDrop">object masterOrMasterShortcutToDrop</param>
        /// <param name="replaceFlags">optional Int32 ReplaceFlags = 0</param>
        [SupportByVersion("Visio", 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVShape ReplaceShape(object masterOrMasterShortcutToDrop, object replaceFlags)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "ReplaceShape", masterOrMasterShortcutToDrop, replaceFlags);
        }

        /// <summary>
        /// SupportByVersion Visio 15,16
        /// </summary>
        /// <param name="masterOrMasterShortcutToDrop">object masterOrMasterShortcutToDrop</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("Visio", 15, 16)]
        public NetOffice.VisioApi.IVShape ReplaceShape(object masterOrMasterShortcutToDrop)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "ReplaceShape", masterOrMasterShortcutToDrop);
        }

        /// <summary>
        /// SupportByVersion Visio 15,16
        /// </summary>
        /// <param name="lineMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices lineMatrix</param>
        /// <param name="fillMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fillMatrix</param>
        /// <param name="effectsMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices effectsMatrix</param>
        /// <param name="fontMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fontMatrix</param>
        /// <param name="lineColor">NetOffice.VisioApi.Enums.VisQuickStyleColors lineColor</param>
        /// <param name="fillColor">NetOffice.VisioApi.Enums.VisQuickStyleColors fillColor</param>
        /// <param name="shadowColor">NetOffice.VisioApi.Enums.VisQuickStyleColors shadowColor</param>
        /// <param name="fontColor">NetOffice.VisioApi.Enums.VisQuickStyleColors fontColor</param>
        [SupportByVersion("Visio", 15, 16)]
        public void SetQuickStyle(NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices lineMatrix, NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fillMatrix, NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices effectsMatrix, NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fontMatrix, NetOffice.VisioApi.Enums.VisQuickStyleColors lineColor, NetOffice.VisioApi.Enums.VisQuickStyleColors fillColor, NetOffice.VisioApi.Enums.VisQuickStyleColors shadowColor, NetOffice.VisioApi.Enums.VisQuickStyleColors fontColor)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetQuickStyle", new object[] { lineMatrix, fillMatrix, effectsMatrix, fontMatrix, lineColor, fillColor, shadowColor, fontColor });
        }

        /// <summary>
        /// SupportByVersion Visio 15,16
        /// </summary>
        /// <param name="fileName">string fileName</param>
        /// <param name="changePictureFlags">optional Int32 ChangePictureFlags = 0</param>
        [SupportByVersion("Visio", 15, 16)]
        public Double ChangePicture(string fileName, object changePictureFlags)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "ChangePicture", fileName, changePictureFlags);
        }

        /// <summary>
        /// SupportByVersion Visio 15,16
        /// </summary>
        /// <param name="fileName">string fileName</param>
        [CustomMethod]
        [SupportByVersion("Visio", 15, 16)]
        public Double ChangePicture(string fileName)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "ChangePicture", fileName);
        }

        #endregion

        #pragma warning restore
    }
}
