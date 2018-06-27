using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// IPageSetup
    /// </summary>
    [SyntaxBypass]
    public class IPageSetup_ : COMObject, NetOffice.ExcelApi.IPageSetup_
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public IPageSetup_() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_PrintQuality(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PrintQuality", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index">optional object index</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void set_PrintQuality(object index, object value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "PrintQuality", index, value);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_PrintQuality
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_PrintQuality")]
        public virtual object PrintQuality(object index)
        {
            return get_PrintQuality(index);
        }

        #endregion

        #region Methods

        #endregion
    }

    /// <summary>
    /// Interface IPageSetup 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class IPageSetup : NetOffice.ExcelApi.Behind.IPageSetup_, NetOffice.ExcelApi.IPageSetup
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
                    _contractType = typeof(NetOffice.ExcelApi.IPageSetup);
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
                    _type = typeof(IPageSetup);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public IPageSetup() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Application Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlCreator Creator
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool BlackAndWhite
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "BlackAndWhite");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BlackAndWhite", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Double BottomMargin
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "BottomMargin");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BottomMargin", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string CenterFooter
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CenterFooter");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CenterFooter", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string CenterHeader
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CenterHeader");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CenterHeader", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool CenterHorizontally
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "CenterHorizontally");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CenterHorizontally", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool CenterVertically
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "CenterVertically");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CenterVertically", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlObjectSize ChartSize
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlObjectSize>(this, "ChartSize");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ChartSize", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool Draft
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Draft");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Draft", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 FirstPageNumber
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "FirstPageNumber");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FirstPageNumber", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object FitToPagesTall
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "FitToPagesTall");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "FitToPagesTall", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object FitToPagesWide
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "FitToPagesWide");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "FitToPagesWide", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Double FooterMargin
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "FooterMargin");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FooterMargin", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Double HeaderMargin
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "HeaderMargin");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HeaderMargin", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string LeftFooter
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "LeftFooter");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LeftFooter", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string LeftHeader
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "LeftHeader");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LeftHeader", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Double LeftMargin
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "LeftMargin");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LeftMargin", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlOrder Order
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlOrder>(this, "Order");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Order", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlPageOrientation Orientation
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlPageOrientation>(this, "Orientation");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Orientation", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlPaperSize PaperSize
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlPaperSize>(this, "PaperSize");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "PaperSize", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string PrintArea
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "PrintArea");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PrintArea", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool PrintGridlines
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PrintGridlines");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PrintGridlines", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool PrintHeadings
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PrintHeadings");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PrintHeadings", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool PrintNotes
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PrintNotes");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PrintNotes", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object PrintQuality
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PrintQuality");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "PrintQuality", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string PrintTitleColumns
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "PrintTitleColumns");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PrintTitleColumns", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string PrintTitleRows
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "PrintTitleRows");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PrintTitleRows", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string RightFooter
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "RightFooter");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RightFooter", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string RightHeader
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "RightHeader");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RightHeader", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Double RightMargin
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "RightMargin");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RightMargin", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Double TopMargin
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "TopMargin");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TopMargin", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Zoom
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Zoom");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Zoom", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlPrintLocation PrintComments
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlPrintLocation>(this, "PrintComments");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "PrintComments", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlPrintErrors PrintErrors
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlPrintErrors>(this, "PrintErrors");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "PrintErrors", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Graphic CenterHeaderPicture
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Graphic>(this, "CenterHeaderPicture", typeof(NetOffice.ExcelApi.Graphic));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Graphic CenterFooterPicture
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Graphic>(this, "CenterFooterPicture", typeof(NetOffice.ExcelApi.Graphic));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Graphic LeftHeaderPicture
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Graphic>(this, "LeftHeaderPicture", typeof(NetOffice.ExcelApi.Graphic));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Graphic LeftFooterPicture
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Graphic>(this, "LeftFooterPicture", typeof(NetOffice.ExcelApi.Graphic));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Graphic RightHeaderPicture
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Graphic>(this, "RightHeaderPicture", typeof(NetOffice.ExcelApi.Graphic));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Graphic RightFooterPicture
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Graphic>(this, "RightFooterPicture", typeof(NetOffice.ExcelApi.Graphic));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool OddAndEvenPagesHeaderFooter
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "OddAndEvenPagesHeaderFooter");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OddAndEvenPagesHeaderFooter", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool DifferentFirstPageHeaderFooter
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DifferentFirstPageHeaderFooter");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DifferentFirstPageHeaderFooter", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool ScaleWithDocHeaderFooter
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ScaleWithDocHeaderFooter");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ScaleWithDocHeaderFooter", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool AlignMarginsHeaderFooter
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AlignMarginsHeaderFooter");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AlignMarginsHeaderFooter", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Pages Pages
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Pages>(this, "Pages", typeof(NetOffice.ExcelApi.Pages));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Page EvenPage
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Page>(this, "EvenPage", typeof(NetOffice.ExcelApi.Page));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Page FirstPage
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Page>(this, "FirstPage", typeof(NetOffice.ExcelApi.Page));
            }
        }

        #endregion

        #region Methods

        #endregion

        #pragma warning restore
    }
}

