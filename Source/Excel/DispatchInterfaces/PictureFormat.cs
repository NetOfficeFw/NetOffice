﻿using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface PictureFormat 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PictureFormat"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
    [Duplicate("NetOffice.OfficeApi.PictureFormat")]
    public class PictureFormat : NetOffice.OfficeApi._IMsoDispObj
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PictureFormat.Parent"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PictureFormat.Brightness"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PictureFormat.ColorType"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PictureFormat.Contrast"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PictureFormat.CropBottom"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PictureFormat.CropLeft"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PictureFormat.CropRight"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PictureFormat.CropTop"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PictureFormat.TransparencyColor"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PictureFormat.TransparentBackground"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PictureFormat.Crop"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.OfficeApi.Crop Crop
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Crop>(this, "Crop", NetOffice.OfficeApi.Crop.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PictureFormat.IncrementBrightness"/> </remarks>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void IncrementBrightness(Single increment)
		{
			 Factory.ExecuteMethod(this, "IncrementBrightness", increment);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PictureFormat.IncrementContrast"/> </remarks>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void IncrementContrast(Single increment)
		{
			 Factory.ExecuteMethod(this, "IncrementContrast", increment);
		}

		#endregion

		#pragma warning restore
	}
}
