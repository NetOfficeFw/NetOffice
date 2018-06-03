using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// Interface IGraphic 
    /// SupportByVersion Excel, 10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class IGraphic : COMObject, NetOffice.ExcelApi.IGraphic
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
                    _type = typeof(IGraphic);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IGraphic() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Application Application
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Enums.XlCreator Creator
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16), ProxyResult]
        public object Parent
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public Single Brightness
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "Brightness");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Brightness", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Enums.MsoPictureColorType ColorType
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPictureColorType>(this, "ColorType");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "ColorType", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public Single Contrast
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "Contrast");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Contrast", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public Single CropBottom
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "CropBottom");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "CropBottom", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public Single CropLeft
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "CropLeft");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "CropLeft", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public Single CropRight
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "CropRight");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "CropRight", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public Single CropTop
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "CropTop");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "CropTop", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public string Filename
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Filename");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Filename", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public Single Height
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "Height");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Height", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Enums.MsoTriState LockAspectRatio
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "LockAspectRatio");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "LockAspectRatio", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public Single Width
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "Width");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Width", value);
            }
        }

        #endregion

        #region Methods

        #endregion

        #pragma warning restore
    }
}

