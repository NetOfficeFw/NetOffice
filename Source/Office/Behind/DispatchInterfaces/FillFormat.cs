using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface FillFormat 
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    [Duplicate("NetOffice.ExcelApi.FillFormat")]
    public class FillFormat : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.FillFormat
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
                    _type = typeof(FillFormat);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public FillFormat() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.ColorFormat BackColor
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ColorFormat>(this, "BackColor", typeof(NetOffice.OfficeApi.ColorFormat));
            }
            set
            {
                Factory.ExecuteReferencePropertySet(this, "BackColor", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.ColorFormat ForeColor
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ColorFormat>(this, "ForeColor", typeof(NetOffice.OfficeApi.ColorFormat));
            }
            set
            {
                Factory.ExecuteReferencePropertySet(this, "ForeColor", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoGradientColorType GradientColorType
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoGradientColorType>(this, "GradientColorType");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Single GradientDegree
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "GradientDegree");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoGradientStyle GradientStyle
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoGradientStyle>(this, "GradientStyle");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 GradientVariant
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "GradientVariant");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoPatternType Pattern
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPatternType>(this, "Pattern");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoPresetGradientType PresetGradientType
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPresetGradientType>(this, "PresetGradientType");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoPresetTexture PresetTexture
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPresetTexture>(this, "PresetTexture");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string TextureName
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "TextureName");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoTextureType TextureType
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTextureType>(this, "TextureType");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Single Transparency
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "Transparency");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Transparency", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoFillType Type
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoFillType>(this, "Type");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoTriState Visible
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Visible");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "Visible", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.GradientStops GradientStops
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.GradientStops>(this, "GradientStops", typeof(NetOffice.OfficeApi.GradientStops));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Single TextureOffsetX
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "TextureOffsetX");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "TextureOffsetX", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Single TextureOffsetY
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "TextureOffsetY");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "TextureOffsetY", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoTextureAlignment TextureAlignment
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTextureAlignment>(this, "TextureAlignment");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "TextureAlignment", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Single TextureHorizontalScale
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "TextureHorizontalScale");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "TextureHorizontalScale", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Single TextureVerticalScale
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "TextureVerticalScale");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "TextureVerticalScale", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoTriState TextureTile
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "TextureTile");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "TextureTile", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoTriState RotateWithObject
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "RotateWithObject");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "RotateWithObject", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual NetOffice.OfficeApi.PictureEffects PictureEffects
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.PictureEffects>(this, "PictureEffects", typeof(NetOffice.OfficeApi.PictureEffects));
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Single GradientAngle
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "GradientAngle");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "GradientAngle", value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Background()
        {
            Factory.ExecuteMethod(this, "Background");
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="style">NetOffice.OfficeApi.Enums.MsoGradientStyle style</param>
        /// <param name="variant">Int32 variant</param>
        /// <param name="degree">Single degree</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void OneColorGradient(NetOffice.OfficeApi.Enums.MsoGradientStyle style, Int32 variant, Single degree)
        {
            Factory.ExecuteMethod(this, "OneColorGradient", style, variant, degree);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pattern">NetOffice.OfficeApi.Enums.MsoPatternType pattern</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Patterned(NetOffice.OfficeApi.Enums.MsoPatternType pattern)
        {
            Factory.ExecuteMethod(this, "Patterned", pattern);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="style">NetOffice.OfficeApi.Enums.MsoGradientStyle style</param>
        /// <param name="variant">Int32 variant</param>
        /// <param name="presetGradientType">NetOffice.OfficeApi.Enums.MsoPresetGradientType presetGradientType</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PresetGradient(NetOffice.OfficeApi.Enums.MsoGradientStyle style, Int32 variant, NetOffice.OfficeApi.Enums.MsoPresetGradientType presetGradientType)
        {
            Factory.ExecuteMethod(this, "PresetGradient", style, variant, presetGradientType);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="presetTexture">NetOffice.OfficeApi.Enums.MsoPresetTexture presetTexture</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PresetTextured(NetOffice.OfficeApi.Enums.MsoPresetTexture presetTexture)
        {
            Factory.ExecuteMethod(this, "PresetTextured", presetTexture);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Solid()
        {
            Factory.ExecuteMethod(this, "Solid");
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="style">NetOffice.OfficeApi.Enums.MsoGradientStyle style</param>
        /// <param name="variant">Int32 variant</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void TwoColorGradient(NetOffice.OfficeApi.Enums.MsoGradientStyle style, Int32 variant)
        {
            Factory.ExecuteMethod(this, "TwoColorGradient", style, variant);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pictureFile">string pictureFile</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void UserPicture(string pictureFile)
        {
            Factory.ExecuteMethod(this, "UserPicture", pictureFile);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="textureFile">string textureFile</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void UserTextured(string textureFile)
        {
            Factory.ExecuteMethod(this, "UserTextured", textureFile);
        }

        #endregion

        #pragma warning restore
    }
}
