using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface FillFormat 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193729.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class FillFormat : COMObject
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

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public FillFormat(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public FillFormat(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FillFormat(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FillFormat(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FillFormat(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FillFormat(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FillFormat() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FillFormat(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197422.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", NetOffice.WordApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191942.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841012.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837539.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ColorFormat BackColor
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ColorFormat>(this, "BackColor", NetOffice.WordApi.ColorFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192378.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ColorFormat ForeColor
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ColorFormat>(this, "ForeColor", NetOffice.WordApi.ColorFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193635.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoGradientColorType GradientColorType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoGradientColorType>(this, "GradientColorType");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838151.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single GradientDegree
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "GradientDegree");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840553.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoGradientStyle GradientStyle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoGradientStyle>(this, "GradientStyle");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838987.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 GradientVariant
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "GradientVariant");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835450.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoPatternType Pattern
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPatternType>(this, "Pattern");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836967.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoPresetGradientType PresetGradientType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPresetGradientType>(this, "PresetGradientType");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823240.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoPresetTexture PresetTexture
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPresetTexture>(this, "PresetTexture");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835189.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string TextureName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TextureName");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195726.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTextureType TextureType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTextureType>(this, "TextureType");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835699.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single Transparency
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836337.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoFillType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoFillType>(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836955.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState Visible
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
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193936.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.OfficeApi.GradientStops GradientStops
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.GradientStops>(this, "GradientStops", NetOffice.OfficeApi.GradientStops.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191863.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public Single TextureOffsetX
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193747.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public Single TextureOffsetY
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837853.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTextureAlignment TextureAlignment
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834560.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public Single TextureHorizontalScale
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191798.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public Single TextureVerticalScale
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197514.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState TextureTile
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
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834590.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState RotateWithObject
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
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821887.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.OfficeApi.PictureEffects PictureEffects
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.PictureEffects>(this, "PictureEffects", NetOffice.OfficeApi.PictureEffects.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835704.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public Single GradientAngle
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Background()
		{
			 Factory.ExecuteMethod(this, "Background");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834854.aspx </remarks>
		/// <param name="style">NetOffice.OfficeApi.Enums.MsoGradientStyle style</param>
		/// <param name="variant">Int32 variant</param>
		/// <param name="degree">Single degree</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void OneColorGradient(NetOffice.OfficeApi.Enums.MsoGradientStyle style, Int32 variant, Single degree)
		{
			 Factory.ExecuteMethod(this, "OneColorGradient", style, variant, degree);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834855.aspx </remarks>
		/// <param name="pattern">NetOffice.OfficeApi.Enums.MsoPatternType pattern</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Patterned(NetOffice.OfficeApi.Enums.MsoPatternType pattern)
		{
			 Factory.ExecuteMethod(this, "Patterned", pattern);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837679.aspx </remarks>
		/// <param name="style">NetOffice.OfficeApi.Enums.MsoGradientStyle style</param>
		/// <param name="variant">Int32 variant</param>
		/// <param name="presetGradientType">NetOffice.OfficeApi.Enums.MsoPresetGradientType presetGradientType</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PresetGradient(NetOffice.OfficeApi.Enums.MsoGradientStyle style, Int32 variant, NetOffice.OfficeApi.Enums.MsoPresetGradientType presetGradientType)
		{
			 Factory.ExecuteMethod(this, "PresetGradient", style, variant, presetGradientType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834881.aspx </remarks>
		/// <param name="presetTexture">NetOffice.OfficeApi.Enums.MsoPresetTexture presetTexture</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PresetTextured(NetOffice.OfficeApi.Enums.MsoPresetTexture presetTexture)
		{
			 Factory.ExecuteMethod(this, "PresetTextured", presetTexture);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193137.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Solid()
		{
			 Factory.ExecuteMethod(this, "Solid");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193717.aspx </remarks>
		/// <param name="style">NetOffice.OfficeApi.Enums.MsoGradientStyle style</param>
		/// <param name="variant">Int32 variant</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void TwoColorGradient(NetOffice.OfficeApi.Enums.MsoGradientStyle style, Int32 variant)
		{
			 Factory.ExecuteMethod(this, "TwoColorGradient", style, variant);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845287.aspx </remarks>
		/// <param name="pictureFile">string pictureFile</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void UserPicture(string pictureFile)
		{
			 Factory.ExecuteMethod(this, "UserPicture", pictureFile);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194653.aspx </remarks>
		/// <param name="textureFile">string textureFile</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void UserTextured(string textureFile)
		{
			 Factory.ExecuteMethod(this, "UserTextured", textureFile);
		}

		#endregion

		#pragma warning restore
	}
}
