using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// DispatchInterface ChInterior 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ChInterior : COMObject, NetOffice.OWC10Api.ChInterior
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
                    _contractType = typeof(NetOffice.OWC10Api.ChInterior);
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
                    _type = typeof(ChInterior);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ChInterior() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual object Color
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Color");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Color", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual object DefaultColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "DefaultColor");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual object BackColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "BackColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "BackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Enums.ChartPatternTypeEnum Pattern
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartPatternTypeEnum>(this, "Pattern");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Enums.ChartFillTypeEnum FillType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartFillTypeEnum>(this, "FillType");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Enums.ChartPresetGradientTypeEnum PresetGradientType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartPresetGradientTypeEnum>(this, "PresetGradientType");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Enums.ChartGradientStyleEnum GradientStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartGradientStyleEnum>(this, "GradientStyle");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Enums.ChartGradientVariantEnum GradientVariant
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartGradientVariantEnum>(this, "GradientVariant");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Double GradientDegree
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "GradientDegree");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Enums.ChartPresetTextureEnum PresetTexture
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartPresetTextureEnum>(this, "PresetTexture");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string TextureName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "TextureName");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Enums.ChartTextureFormatEnum TextureFormat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartTextureFormatEnum>(this, "TextureFormat");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Double TextureStackUnit
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "TextureStackUnit");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Enums.ChartTexturePlacementEnum TexturePlacement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartTexturePlacementEnum>(this, "TexturePlacement");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="patternType">NetOffice.OWC10Api.Enums.ChartPatternTypeEnum patternType</param>
		/// <param name="color">optional object color</param>
		/// <param name="backColor">optional object backColor</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void SetPatterned(NetOffice.OWC10Api.Enums.ChartPatternTypeEnum patternType, object color, object backColor)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetPatterned", patternType, color, backColor);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="patternType">NetOffice.OWC10Api.Enums.ChartPatternTypeEnum patternType</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void SetPatterned(NetOffice.OWC10Api.Enums.ChartPatternTypeEnum patternType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetPatterned", patternType);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="patternType">NetOffice.OWC10Api.Enums.ChartPatternTypeEnum patternType</param>
		/// <param name="color">optional object color</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void SetPatterned(NetOffice.OWC10Api.Enums.ChartPatternTypeEnum patternType, object color)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetPatterned", patternType, color);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="gradientStyle">NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle</param>
		/// <param name="gradientVarient">NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVarient</param>
		/// <param name="gradientPreset">NetOffice.OWC10Api.Enums.ChartPresetGradientTypeEnum gradientPreset</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void SetPresetGradient(NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle, NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVarient, NetOffice.OWC10Api.Enums.ChartPresetGradientTypeEnum gradientPreset)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetPresetGradient", gradientStyle, gradientVarient, gradientPreset);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="textureFile">object textureFile</param>
		/// <param name="textureFormat">optional NetOffice.OWC10Api.Enums.ChartTextureFormatEnum TextureFormat = 4</param>
		/// <param name="stackUnit">optional Double stackUnit = 0</param>
		/// <param name="texturePlacement">optional NetOffice.OWC10Api.Enums.ChartTexturePlacementEnum TexturePlacement = 7</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void SetTextured(object textureFile, object textureFormat, object stackUnit, object texturePlacement)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetTextured", textureFile, textureFormat, stackUnit, texturePlacement);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="textureFile">object textureFile</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void SetTextured(object textureFile)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetTextured", textureFile);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="textureFile">object textureFile</param>
		/// <param name="textureFormat">optional NetOffice.OWC10Api.Enums.ChartTextureFormatEnum TextureFormat = 4</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void SetTextured(object textureFile, object textureFormat)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetTextured", textureFile, textureFormat);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="textureFile">object textureFile</param>
		/// <param name="textureFormat">optional NetOffice.OWC10Api.Enums.ChartTextureFormatEnum TextureFormat = 4</param>
		/// <param name="stackUnit">optional Double stackUnit = 0</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void SetTextured(object textureFile, object textureFormat, object stackUnit)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetTextured", textureFile, textureFormat, stackUnit);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="gradientStyle">NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle</param>
		/// <param name="gradientVariant">NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant</param>
		/// <param name="gradientDegree">Double gradientDegree</param>
		/// <param name="color">optional object color</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void SetOneColorGradient(NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle, NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant, Double gradientDegree, object color)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetOneColorGradient", gradientStyle, gradientVariant, gradientDegree, color);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="gradientStyle">NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle</param>
		/// <param name="gradientVariant">NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant</param>
		/// <param name="gradientDegree">Double gradientDegree</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void SetOneColorGradient(NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle, NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant, Double gradientDegree)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetOneColorGradient", gradientStyle, gradientVariant, gradientDegree);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="color">optional object color</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void SetSolid(object color)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetSolid", color);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void SetSolid()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetSolid");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="gradientStyle">NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle</param>
		/// <param name="gradientVariant">NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant</param>
		/// <param name="color">optional object color</param>
		/// <param name="backColor">optional object backColor</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void SetTwoColorGradient(NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle, NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant, object color, object backColor)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetTwoColorGradient", gradientStyle, gradientVariant, color, backColor);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="gradientStyle">NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle</param>
		/// <param name="gradientVariant">NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void SetTwoColorGradient(NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle, NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetTwoColorGradient", gradientStyle, gradientVariant);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="gradientStyle">NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle</param>
		/// <param name="gradientVariant">NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant</param>
		/// <param name="color">optional object color</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void SetTwoColorGradient(NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle, NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant, object color)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetTwoColorGradient", gradientStyle, gradientVariant, color);
		}

		#endregion

		#pragma warning restore
	}
}

