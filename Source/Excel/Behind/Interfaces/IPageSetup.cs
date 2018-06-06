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
            return Factory.ExecuteVariantPropertyGet(this, "PrintQuality", index);
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
            Factory.ExecutePropertySet(this, "PrintQuality", index, value);
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
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
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
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
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
                return Factory.ExecuteReferencePropertyGet(this, "Parent");
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
                return Factory.ExecuteBoolPropertyGet(this, "BlackAndWhite");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "BlackAndWhite", value);
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
                return Factory.ExecuteDoublePropertyGet(this, "BottomMargin");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "BottomMargin", value);
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
                return Factory.ExecuteStringPropertyGet(this, "CenterFooter");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "CenterFooter", value);
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
                return Factory.ExecuteStringPropertyGet(this, "CenterHeader");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "CenterHeader", value);
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
                return Factory.ExecuteBoolPropertyGet(this, "CenterHorizontally");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "CenterHorizontally", value);
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
                return Factory.ExecuteBoolPropertyGet(this, "CenterVertically");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "CenterVertically", value);
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
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlObjectSize>(this, "ChartSize");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "ChartSize", value);
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
                return Factory.ExecuteBoolPropertyGet(this, "Draft");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Draft", value);
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
                return Factory.ExecuteInt32PropertyGet(this, "FirstPageNumber");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "FirstPageNumber", value);
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
                return Factory.ExecuteVariantPropertyGet(this, "FitToPagesTall");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "FitToPagesTall", value);
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
                return Factory.ExecuteVariantPropertyGet(this, "FitToPagesWide");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "FitToPagesWide", value);
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
                return Factory.ExecuteDoublePropertyGet(this, "FooterMargin");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "FooterMargin", value);
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
                return Factory.ExecuteDoublePropertyGet(this, "HeaderMargin");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "HeaderMargin", value);
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
                return Factory.ExecuteStringPropertyGet(this, "LeftFooter");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "LeftFooter", value);
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
                return Factory.ExecuteStringPropertyGet(this, "LeftHeader");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "LeftHeader", value);
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
                return Factory.ExecuteDoublePropertyGet(this, "LeftMargin");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "LeftMargin", value);
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
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlOrder>(this, "Order");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "Order", value);
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
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlPageOrientation>(this, "Orientation");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "Orientation", value);
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
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlPaperSize>(this, "PaperSize");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "PaperSize", value);
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
                return Factory.ExecuteStringPropertyGet(this, "PrintArea");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "PrintArea", value);
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
                return Factory.ExecuteBoolPropertyGet(this, "PrintGridlines");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "PrintGridlines", value);
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
                return Factory.ExecuteBoolPropertyGet(this, "PrintHeadings");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "PrintHeadings", value);
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
                return Factory.ExecuteBoolPropertyGet(this, "PrintNotes");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "PrintNotes", value);
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
                return Factory.ExecuteVariantPropertyGet(this, "PrintQuality");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "PrintQuality", value);
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
                return Factory.ExecuteStringPropertyGet(this, "PrintTitleColumns");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "PrintTitleColumns", value);
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
                return Factory.ExecuteStringPropertyGet(this, "PrintTitleRows");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "PrintTitleRows", value);
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
                return Factory.ExecuteStringPropertyGet(this, "RightFooter");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "RightFooter", value);
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
                return Factory.ExecuteStringPropertyGet(this, "RightHeader");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "RightHeader", value);
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
                return Factory.ExecuteDoublePropertyGet(this, "RightMargin");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "RightMargin", value);
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
                return Factory.ExecuteDoublePropertyGet(this, "TopMargin");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "TopMargin", value);
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
                return Factory.ExecuteVariantPropertyGet(this, "Zoom");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "Zoom", value);
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
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlPrintLocation>(this, "PrintComments");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "PrintComments", value);
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
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlPrintErrors>(this, "PrintErrors");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "PrintErrors", value);
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
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Graphic>(this, "CenterHeaderPicture", typeof(NetOffice.ExcelApi.Graphic));
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
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Graphic>(this, "CenterFooterPicture", typeof(NetOffice.ExcelApi.Graphic));
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
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Graphic>(this, "LeftHeaderPicture", typeof(NetOffice.ExcelApi.Graphic));
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
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Graphic>(this, "LeftFooterPicture", typeof(NetOffice.ExcelApi.Graphic));
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
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Graphic>(this, "RightHeaderPicture", typeof(NetOffice.ExcelApi.Graphic));
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
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Graphic>(this, "RightFooterPicture", typeof(NetOffice.ExcelApi.Graphic));
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
                return Factory.ExecuteBoolPropertyGet(this, "OddAndEvenPagesHeaderFooter");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "OddAndEvenPagesHeaderFooter", value);
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
                return Factory.ExecuteBoolPropertyGet(this, "DifferentFirstPageHeaderFooter");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "DifferentFirstPageHeaderFooter", value);
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
                return Factory.ExecuteBoolPropertyGet(this, "ScaleWithDocHeaderFooter");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "ScaleWithDocHeaderFooter", value);
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
                return Factory.ExecuteBoolPropertyGet(this, "AlignMarginsHeaderFooter");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "AlignMarginsHeaderFooter", value);
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
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Pages>(this, "Pages", typeof(NetOffice.ExcelApi.Pages));
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
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Page>(this, "EvenPage", typeof(NetOffice.ExcelApi.Page));
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
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Page>(this, "FirstPage", typeof(NetOffice.ExcelApi.Page));
            }
        }

        #endregion

        #region Methods

        #endregion

        #pragma warning restore
    }
}

