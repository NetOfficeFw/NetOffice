﻿using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// Interface IFont 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class IFont : COMObject, NetOffice.ExcelApi.IFont
    {
        #pragma warning disable

        #region Type Information

        /// <summary>        /// Instance Type
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
                    _type = typeof(IFont);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IFont() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Application Application
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
        public NetOffice.ExcelApi.Enums.XlCreator Creator
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
        public object Parent
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
        public object Background
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "Background");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "Background", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object Bold
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "Bold");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "Bold", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object Color
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "Color");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "Color", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object ColorIndex
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "ColorIndex");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "ColorIndex", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object FontStyle
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "FontStyle");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "FontStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object Italic
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "Italic");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "Italic", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object Name
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "Name");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "Name", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object OutlineFont
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "OutlineFont");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "OutlineFont", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object Shadow
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "Shadow");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "Shadow", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object Size
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "Size");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "Size", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object Strikethrough
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "Strikethrough");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "Strikethrough", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object Subscript
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "Subscript");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "Subscript", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object Superscript
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "Superscript");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "Superscript", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object Underline
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "Underline");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "Underline", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public object ThemeColor
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "ThemeColor");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "ThemeColor", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public object TintAndShade
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "TintAndShade");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "TintAndShade", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Enums.XlThemeFont ThemeFont
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlThemeFont>(this, "ThemeFont");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "ThemeFont", value);
            }
        }

        #endregion

        #region Methods

        #endregion

        #pragma warning restore
    }
}
