using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface ChInterior 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ChInterior : COMObject
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
                    _type = typeof(ChInterior);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ChInterior(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ChInterior(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChInterior(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChInterior(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChInterior(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChInterior(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChInterior() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChInterior(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
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
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object DefaultColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DefaultColor");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object BackColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "BackColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "BackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartPatternTypeEnum Pattern
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartPatternTypeEnum>(this, "Pattern");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartFillTypeEnum FillType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartFillTypeEnum>(this, "FillType");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartPresetGradientTypeEnum PresetGradientType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartPresetGradientTypeEnum>(this, "PresetGradientType");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartGradientStyleEnum GradientStyle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartGradientStyleEnum>(this, "GradientStyle");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartGradientVariantEnum GradientVariant
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartGradientVariantEnum>(this, "GradientVariant");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Double GradientDegree
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "GradientDegree");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartPresetTextureEnum PresetTexture
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartPresetTextureEnum>(this, "PresetTexture");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string TextureName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TextureName");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartTextureFormatEnum TextureFormat
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartTextureFormatEnum>(this, "TextureFormat");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Double TextureStackUnit
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "TextureStackUnit");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartTexturePlacementEnum TexturePlacement
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartTexturePlacementEnum>(this, "TexturePlacement");
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
		public void SetPatterned(NetOffice.OWC10Api.Enums.ChartPatternTypeEnum patternType, object color, object backColor)
		{
			 Factory.ExecuteMethod(this, "SetPatterned", patternType, color, backColor);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="patternType">NetOffice.OWC10Api.Enums.ChartPatternTypeEnum patternType</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void SetPatterned(NetOffice.OWC10Api.Enums.ChartPatternTypeEnum patternType)
		{
			 Factory.ExecuteMethod(this, "SetPatterned", patternType);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="patternType">NetOffice.OWC10Api.Enums.ChartPatternTypeEnum patternType</param>
		/// <param name="color">optional object color</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void SetPatterned(NetOffice.OWC10Api.Enums.ChartPatternTypeEnum patternType, object color)
		{
			 Factory.ExecuteMethod(this, "SetPatterned", patternType, color);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="gradientStyle">NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle</param>
		/// <param name="gradientVarient">NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVarient</param>
		/// <param name="gradientPreset">NetOffice.OWC10Api.Enums.ChartPresetGradientTypeEnum gradientPreset</param>
		[SupportByVersion("OWC10", 1)]
		public void SetPresetGradient(NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle, NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVarient, NetOffice.OWC10Api.Enums.ChartPresetGradientTypeEnum gradientPreset)
		{
			 Factory.ExecuteMethod(this, "SetPresetGradient", gradientStyle, gradientVarient, gradientPreset);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="textureFile">object textureFile</param>
		/// <param name="textureFormat">optional NetOffice.OWC10Api.Enums.ChartTextureFormatEnum TextureFormat = 4</param>
		/// <param name="stackUnit">optional Double stackUnit = 0</param>
		/// <param name="texturePlacement">optional NetOffice.OWC10Api.Enums.ChartTexturePlacementEnum TexturePlacement = 7</param>
		[SupportByVersion("OWC10", 1)]
		public void SetTextured(object textureFile, object textureFormat, object stackUnit, object texturePlacement)
		{
			 Factory.ExecuteMethod(this, "SetTextured", textureFile, textureFormat, stackUnit, texturePlacement);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="textureFile">object textureFile</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void SetTextured(object textureFile)
		{
			 Factory.ExecuteMethod(this, "SetTextured", textureFile);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="textureFile">object textureFile</param>
		/// <param name="textureFormat">optional NetOffice.OWC10Api.Enums.ChartTextureFormatEnum TextureFormat = 4</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void SetTextured(object textureFile, object textureFormat)
		{
			 Factory.ExecuteMethod(this, "SetTextured", textureFile, textureFormat);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="textureFile">object textureFile</param>
		/// <param name="textureFormat">optional NetOffice.OWC10Api.Enums.ChartTextureFormatEnum TextureFormat = 4</param>
		/// <param name="stackUnit">optional Double stackUnit = 0</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void SetTextured(object textureFile, object textureFormat, object stackUnit)
		{
			 Factory.ExecuteMethod(this, "SetTextured", textureFile, textureFormat, stackUnit);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="gradientStyle">NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle</param>
		/// <param name="gradientVariant">NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant</param>
		/// <param name="gradientDegree">Double gradientDegree</param>
		/// <param name="color">optional object color</param>
		[SupportByVersion("OWC10", 1)]
		public void SetOneColorGradient(NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle, NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant, Double gradientDegree, object color)
		{
			 Factory.ExecuteMethod(this, "SetOneColorGradient", gradientStyle, gradientVariant, gradientDegree, color);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="gradientStyle">NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle</param>
		/// <param name="gradientVariant">NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant</param>
		/// <param name="gradientDegree">Double gradientDegree</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void SetOneColorGradient(NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle, NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant, Double gradientDegree)
		{
			 Factory.ExecuteMethod(this, "SetOneColorGradient", gradientStyle, gradientVariant, gradientDegree);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="color">optional object color</param>
		[SupportByVersion("OWC10", 1)]
		public void SetSolid(object color)
		{
			 Factory.ExecuteMethod(this, "SetSolid", color);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void SetSolid()
		{
			 Factory.ExecuteMethod(this, "SetSolid");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="gradientStyle">NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle</param>
		/// <param name="gradientVariant">NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant</param>
		/// <param name="color">optional object color</param>
		/// <param name="backColor">optional object backColor</param>
		[SupportByVersion("OWC10", 1)]
		public void SetTwoColorGradient(NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle, NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant, object color, object backColor)
		{
			 Factory.ExecuteMethod(this, "SetTwoColorGradient", gradientStyle, gradientVariant, color, backColor);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="gradientStyle">NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle</param>
		/// <param name="gradientVariant">NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void SetTwoColorGradient(NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle, NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant)
		{
			 Factory.ExecuteMethod(this, "SetTwoColorGradient", gradientStyle, gradientVariant);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="gradientStyle">NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle</param>
		/// <param name="gradientVariant">NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant</param>
		/// <param name="color">optional object color</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void SetTwoColorGradient(NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle, NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant, object color)
		{
			 Factory.ExecuteMethod(this, "SetTwoColorGradient", gradientStyle, gradientVariant, color);
		}

		#endregion

		#pragma warning restore
	}
}
