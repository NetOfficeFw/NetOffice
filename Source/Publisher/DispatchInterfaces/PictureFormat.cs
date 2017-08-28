using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface PictureFormat 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class PictureFormat : COMObject
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
                    _type = typeof(PictureFormat);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public PictureFormat(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PictureFormat(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PictureFormat(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PictureFormat(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PictureFormat(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PictureFormat(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PictureFormat() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PictureFormat(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Application>(this, "Application", NetOffice.PublisherApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
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
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
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
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
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
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public object CropBottom
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "CropBottom");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "CropBottom", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public object CropLeft
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "CropLeft");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "CropLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public object CropRight
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "CropRight");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "CropRight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public object CropTop
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "CropTop");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "CropTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 TransparencyColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "TransparencyColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TransparencyColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState TransparentBackground
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "TransparentBackground");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TransparentBackground", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbVerticalPictureLocking VerticalPictureLocking
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbVerticalPictureLocking>(this, "VerticalPictureLocking");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "VerticalPictureLocking", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbHorizontalPictureLocking HorizontalPictureLocking
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbHorizontalPictureLocking>(this, "HorizontalPictureLocking");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "HorizontalPictureLocking", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbColorModel ColorModel
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbColorModel>(this, "ColorModel");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 ColorsInPalette
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ColorsInPalette");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 EffectiveResolution
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "EffectiveResolution");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public string Filename
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Filename");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 FileSize
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "FileSize");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState HasAlphaChannel
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "HasAlphaChannel");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public object Height
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Height");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 HorizontalScale
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "HorizontalScale");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbImageFormat ImageFormat
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbImageFormat>(this, "ImageFormat");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState IsGreyScale
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "IsGreyScale");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState IsLinked
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "IsLinked");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState IsTrueColor
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "IsTrueColor");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 OriginalColorsInPalette
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "OriginalColorsInPalette");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 OriginalFileSize
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "OriginalFileSize");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState OriginalHasAlphaChannel
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "OriginalHasAlphaChannel");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public object OriginalHeight
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "OriginalHeight");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState OriginalIsTrueColor
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "OriginalIsTrueColor");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 OriginalResolution
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "OriginalResolution");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public object OriginalWidth
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "OriginalWidth");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 VerticalScale
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "VerticalScale");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public object Width
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Width");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbLinkedFileStatus LinkedFileStatus
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbLinkedFileStatus>(this, "LinkedFileStatus");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState IsEmpty
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "IsEmpty");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool HasTransparencyColor
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasTransparencyColor");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ColorFormat RecoloredPictureColor
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ColorFormat>(this, "RecoloredPictureColor", NetOffice.PublisherApi.ColorFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState IsRecolored
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "IsRecolored");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState LeaveBlackAsBlack
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "LeaveBlackAsBlack");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void IncrementBrightness(Single increment)
		{
			 Factory.ExecuteMethod(this, "IncrementBrightness", increment);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void IncrementContrast(Single increment)
		{
			 Factory.ExecuteMethod(this, "IncrementContrast", increment);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="pathname">string pathname</param>
		/// <param name="insertAs">optional NetOffice.PublisherApi.Enums.PbPictureInsertAs InsertAs = 3</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void Replace(string pathname, object insertAs)
		{
			 Factory.ExecuteMethod(this, "Replace", pathname, insertAs);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="pathname">string pathname</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void Replace(string pathname)
		{
			 Factory.ExecuteMethod(this, "Replace", pathname);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="color">NetOffice.PublisherApi.ColorFormat color</param>
		/// <param name="leaveBlackPartsBlack">NetOffice.OfficeApi.Enums.MsoTriState leaveBlackPartsBlack</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void Recolor(NetOffice.PublisherApi.ColorFormat color, NetOffice.OfficeApi.Enums.MsoTriState leaveBlackPartsBlack)
		{
			 Factory.ExecuteMethod(this, "Recolor", color, leaveBlackPartsBlack);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void RestoreOriginalColors()
		{
			 Factory.ExecuteMethod(this, "RestoreOriginalColors");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void FillFrame()
		{
			 Factory.ExecuteMethod(this, "FillFrame");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void FitFrame()
		{
			 Factory.ExecuteMethod(this, "FitFrame");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void ClearCrop()
		{
			 Factory.ExecuteMethod(this, "ClearCrop");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void Remove()
		{
			 Factory.ExecuteMethod(this, "Remove");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="pathname">string pathname</param>
		/// <param name="insertAs">optional NetOffice.PublisherApi.Enums.PbPictureInsertAs InsertAs = 3</param>
		/// <param name="fit">optional NetOffice.PublisherApi.Enums.pbPictureInsertFit Fit = 1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void ReplaceEx(string pathname, object insertAs, object fit)
		{
			 Factory.ExecuteMethod(this, "ReplaceEx", pathname, insertAs, fit);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="pathname">string pathname</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ReplaceEx(string pathname)
		{
			 Factory.ExecuteMethod(this, "ReplaceEx", pathname);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="pathname">string pathname</param>
		/// <param name="insertAs">optional NetOffice.PublisherApi.Enums.PbPictureInsertAs InsertAs = 3</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ReplaceEx(string pathname, object insertAs)
		{
			 Factory.ExecuteMethod(this, "ReplaceEx", pathname, insertAs);
		}

		#endregion

		#pragma warning restore
	}
}
