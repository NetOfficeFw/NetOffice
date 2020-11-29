using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface FillFormat 
	/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat"/> </remarks>
	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.Application"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16), ProxyResult]
		public object Application
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.Creator"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.Parent"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.BackColor"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.ColorFormat BackColor
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ColorFormat>(this, "BackColor", NetOffice.PowerPointApi.ColorFormat.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "BackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.ForeColor"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.ColorFormat ForeColor
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ColorFormat>(this, "ForeColor", NetOffice.PowerPointApi.ColorFormat.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "ForeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.GradientColorType"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoGradientColorType GradientColorType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoGradientColorType>(this, "GradientColorType");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.GradientDegree"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Single GradientDegree
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "GradientDegree");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.GradientStyle"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoGradientStyle GradientStyle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoGradientStyle>(this, "GradientStyle");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.GradientVariant"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Int32 GradientVariant
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "GradientVariant");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.Pattern"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoPatternType Pattern
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPatternType>(this, "Pattern");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.PresetGradientType"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoPresetGradientType PresetGradientType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPresetGradientType>(this, "PresetGradientType");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.PresetTexture"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoPresetTexture PresetTexture
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPresetTexture>(this, "PresetTexture");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.TextureName"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public string TextureName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TextureName");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.TextureType"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTextureType TextureType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTextureType>(this, "TextureType");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.Transparency"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.Type"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoFillType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoFillType>(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.Visible"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.GradientStops"/> </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.GradientStops GradientStops
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.GradientStops>(this, "GradientStops", NetOffice.OfficeApi.GradientStops.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.TextureOffsetX"/> </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
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
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.TextureOffsetY"/> </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
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
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.TextureAlignment"/> </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
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
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.TextureHorizontalScale"/> </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
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
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.TextureVerticalScale"/> </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
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
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.TextureTile"/> </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
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
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.RotateWithObject"/> </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.PictureEffects"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.OfficeApi.PictureEffects PictureEffects
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.PictureEffects>(this, "PictureEffects", NetOffice.OfficeApi.PictureEffects.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.GradientAngle"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.Background"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Background()
		{
			 Factory.ExecuteMethod(this, "Background");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.OneColorGradient"/> </remarks>
		/// <param name="style">NetOffice.OfficeApi.Enums.MsoGradientStyle style</param>
		/// <param name="variant">Int32 variant</param>
		/// <param name="degree">Single degree</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void OneColorGradient(NetOffice.OfficeApi.Enums.MsoGradientStyle style, Int32 variant, Single degree)
		{
			 Factory.ExecuteMethod(this, "OneColorGradient", style, variant, degree);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.Patterned"/> </remarks>
		/// <param name="pattern">NetOffice.OfficeApi.Enums.MsoPatternType pattern</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Patterned(NetOffice.OfficeApi.Enums.MsoPatternType pattern)
		{
			 Factory.ExecuteMethod(this, "Patterned", pattern);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.PresetGradient"/> </remarks>
		/// <param name="style">NetOffice.OfficeApi.Enums.MsoGradientStyle style</param>
		/// <param name="variant">Int32 variant</param>
		/// <param name="presetGradientType">NetOffice.OfficeApi.Enums.MsoPresetGradientType presetGradientType</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void PresetGradient(NetOffice.OfficeApi.Enums.MsoGradientStyle style, Int32 variant, NetOffice.OfficeApi.Enums.MsoPresetGradientType presetGradientType)
		{
			 Factory.ExecuteMethod(this, "PresetGradient", style, variant, presetGradientType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.PresetTextured"/> </remarks>
		/// <param name="presetTexture">NetOffice.OfficeApi.Enums.MsoPresetTexture presetTexture</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void PresetTextured(NetOffice.OfficeApi.Enums.MsoPresetTexture presetTexture)
		{
			 Factory.ExecuteMethod(this, "PresetTextured", presetTexture);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.Solid"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Solid()
		{
			 Factory.ExecuteMethod(this, "Solid");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.TwoColorGradient"/> </remarks>
		/// <param name="style">NetOffice.OfficeApi.Enums.MsoGradientStyle style</param>
		/// <param name="variant">Int32 variant</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void TwoColorGradient(NetOffice.OfficeApi.Enums.MsoGradientStyle style, Int32 variant)
		{
			 Factory.ExecuteMethod(this, "TwoColorGradient", style, variant);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.UserPicture"/> </remarks>
		/// <param name="pictureFile">string pictureFile</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void UserPicture(string pictureFile)
		{
			 Factory.ExecuteMethod(this, "UserPicture", pictureFile);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.FillFormat.UserTextured"/> </remarks>
		/// <param name="textureFile">string textureFile</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void UserTextured(string textureFile)
		{
			 Factory.ExecuteMethod(this, "UserTextured", textureFile);
		}

		#endregion

		#pragma warning restore
	}
}
