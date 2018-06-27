using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface ChartFillFormat 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class ChartFillFormat : COMObject, NetOffice.OfficeApi.ChartFillFormat
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
                    _contractType = typeof(NetOffice.OfficeApi.ChartFillFormat);
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
                    _type = typeof(ChartFillFormat);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ChartFillFormat() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.ChartColorFormat BackColor
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ChartColorFormat>(this, "BackColor", typeof(NetOffice.OfficeApi.ChartColorFormat));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.ChartColorFormat ForeColor
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ChartColorFormat>(this, "ForeColor", typeof(NetOffice.OfficeApi.ChartColorFormat));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 GradientColorType
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "GradientColorType");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Single GradientDegree
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "GradientDegree");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 GradientStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "GradientStyle");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 GradientVariant
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "GradientVariant");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 Pattern
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Pattern");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 PresetGradientType
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "PresetGradientType");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 PresetTexture
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "PresetTexture");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string TextureName
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "TextureName");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 TextureType
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "TextureType");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 Type
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Type");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 Visible
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Visible");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Visible", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16), ProxyResult]
        public virtual object Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Application");
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Int32 Creator
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="style">Int32 style</param>
        /// <param name="variant">Int32 variant</param>
        /// <param name="degree">Single degree</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void OneColorGradient(Int32 style, Int32 variant, Single degree)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "OneColorGradient", style, variant, degree);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="style">Int32 style</param>
        /// <param name="variant">Int32 variant</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void TwoColorGradient(Int32 style, Int32 variant)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "TwoColorGradient", style, variant);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="presetTexture">Int32 presetTexture</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void PresetTextured(Int32 presetTexture)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PresetTextured", presetTexture);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void Solid()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Solid");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="pattern">Int32 pattern</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void Patterned(Int32 pattern)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Patterned", pattern);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="pictureFile">object pictureFile</param>
        /// <param name="pictureFormat">object pictureFormat</param>
        /// <param name="pictureStackUnit">object pictureStackUnit</param>
        /// <param name="picturePlacement">object picturePlacement</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void UserPicture(object pictureFile, object pictureFormat, object pictureStackUnit, object picturePlacement)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "UserPicture", pictureFile, pictureFormat, pictureStackUnit, picturePlacement);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="textureFile">string textureFile</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void UserTextured(string textureFile)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "UserTextured", textureFile);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="style">Int32 style</param>
        /// <param name="variant">Int32 variant</param>
        /// <param name="presetGradientType">Int32 presetGradientType</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void PresetGradient(Int32 style, Int32 variant, Int32 presetGradientType)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PresetGradient", style, variant, presetGradientType);
        }

        #endregion

        #pragma warning restore
    }
}
