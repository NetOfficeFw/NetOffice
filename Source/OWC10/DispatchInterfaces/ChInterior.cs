using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OWC10Api
{
	///<summary>
	/// DispatchInterface ChInterior 
	/// SupportByVersion OWC10, 1
	///</summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class ChInterior : COMObject
	{
		#pragma warning disable
		#region Type Information

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
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		[SupportByVersionAttribute("OWC10", 1)]
		public object Color
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Color", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Color", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public object DefaultColor
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultColor", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public object BackColor
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BackColor", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "BackColor", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartPatternTypeEnum Pattern
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Pattern", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OWC10Api.Enums.ChartPatternTypeEnum)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartFillTypeEnum FillType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FillType", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OWC10Api.Enums.ChartFillTypeEnum)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartPresetGradientTypeEnum PresetGradientType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PresetGradientType", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OWC10Api.Enums.ChartPresetGradientTypeEnum)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartGradientStyleEnum GradientStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GradientStyle", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OWC10Api.Enums.ChartGradientStyleEnum)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartGradientVariantEnum GradientVariant
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GradientVariant", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OWC10Api.Enums.ChartGradientVariantEnum)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public Double GradientDegree
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GradientDegree", paramsArray);
				return NetRuntimeSystem.Convert.ToDouble(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartPresetTextureEnum PresetTexture
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PresetTexture", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OWC10Api.Enums.ChartPresetTextureEnum)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public string TextureName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TextureName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartTextureFormatEnum TextureFormat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TextureFormat", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OWC10Api.Enums.ChartTextureFormatEnum)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public Double TextureStackUnit
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TextureStackUnit", paramsArray);
				return NetRuntimeSystem.Convert.ToDouble(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartTexturePlacementEnum TexturePlacement
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TexturePlacement", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OWC10Api.Enums.ChartTexturePlacementEnum)intReturnItem;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="patternType">NetOffice.OWC10Api.Enums.ChartPatternTypeEnum patternType</param>
		/// <param name="color">optional object Color</param>
		/// <param name="backColor">optional object BackColor</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void SetPatterned(NetOffice.OWC10Api.Enums.ChartPatternTypeEnum patternType, object color, object backColor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(patternType, color, backColor);
			Invoker.Method(this, "SetPatterned", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="patternType">NetOffice.OWC10Api.Enums.ChartPatternTypeEnum patternType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void SetPatterned(NetOffice.OWC10Api.Enums.ChartPatternTypeEnum patternType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(patternType);
			Invoker.Method(this, "SetPatterned", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="patternType">NetOffice.OWC10Api.Enums.ChartPatternTypeEnum patternType</param>
		/// <param name="color">optional object Color</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void SetPatterned(NetOffice.OWC10Api.Enums.ChartPatternTypeEnum patternType, object color)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(patternType, color);
			Invoker.Method(this, "SetPatterned", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="gradientStyle">NetOffice.OWC10Api.Enums.ChartGradientStyleEnum GradientStyle</param>
		/// <param name="gradientVarient">NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVarient</param>
		/// <param name="gradientPreset">NetOffice.OWC10Api.Enums.ChartPresetGradientTypeEnum gradientPreset</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void SetPresetGradient(NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle, NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVarient, NetOffice.OWC10Api.Enums.ChartPresetGradientTypeEnum gradientPreset)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(gradientStyle, gradientVarient, gradientPreset);
			Invoker.Method(this, "SetPresetGradient", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="textureFile">object textureFile</param>
		/// <param name="textureFormat">optional NetOffice.OWC10Api.Enums.ChartTextureFormatEnum TextureFormat = 4</param>
		/// <param name="stackUnit">optional Double stackUnit = 0</param>
		/// <param name="texturePlacement">optional NetOffice.OWC10Api.Enums.ChartTexturePlacementEnum TexturePlacement = 7</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void SetTextured(object textureFile, object textureFormat, object stackUnit, object texturePlacement)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(textureFile, textureFormat, stackUnit, texturePlacement);
			Invoker.Method(this, "SetTextured", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="textureFile">object textureFile</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void SetTextured(object textureFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(textureFile);
			Invoker.Method(this, "SetTextured", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="textureFile">object textureFile</param>
		/// <param name="textureFormat">optional NetOffice.OWC10Api.Enums.ChartTextureFormatEnum TextureFormat = 4</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void SetTextured(object textureFile, object textureFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(textureFile, textureFormat);
			Invoker.Method(this, "SetTextured", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="textureFile">object textureFile</param>
		/// <param name="textureFormat">optional NetOffice.OWC10Api.Enums.ChartTextureFormatEnum TextureFormat = 4</param>
		/// <param name="stackUnit">optional Double stackUnit = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void SetTextured(object textureFile, object textureFormat, object stackUnit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(textureFile, textureFormat, stackUnit);
			Invoker.Method(this, "SetTextured", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="gradientStyle">NetOffice.OWC10Api.Enums.ChartGradientStyleEnum GradientStyle</param>
		/// <param name="gradientVariant">NetOffice.OWC10Api.Enums.ChartGradientVariantEnum GradientVariant</param>
		/// <param name="gradientDegree">Double GradientDegree</param>
		/// <param name="color">optional object Color</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void SetOneColorGradient(NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle, NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant, Double gradientDegree, object color)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(gradientStyle, gradientVariant, gradientDegree, color);
			Invoker.Method(this, "SetOneColorGradient", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="gradientStyle">NetOffice.OWC10Api.Enums.ChartGradientStyleEnum GradientStyle</param>
		/// <param name="gradientVariant">NetOffice.OWC10Api.Enums.ChartGradientVariantEnum GradientVariant</param>
		/// <param name="gradientDegree">Double GradientDegree</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void SetOneColorGradient(NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle, NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant, Double gradientDegree)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(gradientStyle, gradientVariant, gradientDegree);
			Invoker.Method(this, "SetOneColorGradient", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="color">optional object Color</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void SetSolid(object color)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(color);
			Invoker.Method(this, "SetSolid", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void SetSolid()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SetSolid", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="gradientStyle">NetOffice.OWC10Api.Enums.ChartGradientStyleEnum GradientStyle</param>
		/// <param name="gradientVariant">NetOffice.OWC10Api.Enums.ChartGradientVariantEnum GradientVariant</param>
		/// <param name="color">optional object Color</param>
		/// <param name="backColor">optional object BackColor</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void SetTwoColorGradient(NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle, NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant, object color, object backColor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(gradientStyle, gradientVariant, color, backColor);
			Invoker.Method(this, "SetTwoColorGradient", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="gradientStyle">NetOffice.OWC10Api.Enums.ChartGradientStyleEnum GradientStyle</param>
		/// <param name="gradientVariant">NetOffice.OWC10Api.Enums.ChartGradientVariantEnum GradientVariant</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void SetTwoColorGradient(NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle, NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(gradientStyle, gradientVariant);
			Invoker.Method(this, "SetTwoColorGradient", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="gradientStyle">NetOffice.OWC10Api.Enums.ChartGradientStyleEnum GradientStyle</param>
		/// <param name="gradientVariant">NetOffice.OWC10Api.Enums.ChartGradientVariantEnum GradientVariant</param>
		/// <param name="color">optional object Color</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void SetTwoColorGradient(NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle, NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant, object color)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(gradientStyle, gradientVariant, color);
			Invoker.Method(this, "SetTwoColorGradient", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}