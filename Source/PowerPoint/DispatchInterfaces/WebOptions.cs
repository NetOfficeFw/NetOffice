using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface WebOptions 
	/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class WebOptions : COMObject
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
                    _type = typeof(WebOptions);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public WebOptions(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public WebOptions(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public WebOptions(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public WebOptions(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public WebOptions(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public WebOptions(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public WebOptions() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public WebOptions(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState IncludeNavigation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "IncludeNavigation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "IncludeNavigation", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpFrameColors FrameColors
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpFrameColors>(this, "FrameColors");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "FrameColors", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState ResizeGraphics
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "ResizeGraphics");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ResizeGraphics", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState ShowSlideAnimation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "ShowSlideAnimation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ShowSlideAnimation", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState OrganizeInFolder
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "OrganizeInFolder");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "OrganizeInFolder", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState UseLongFileNames
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "UseLongFileNames");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "UseLongFileNames", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState RelyOnVML
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "RelyOnVML");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RelyOnVML", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState AllowPNG
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "AllowPNG");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "AllowPNG", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoScreenSize ScreenSize
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoScreenSize>(this, "ScreenSize");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ScreenSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoEncoding Encoding
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoEncoding>(this, "Encoding");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Encoding", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public string FolderSuffix
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FolderSuffix");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTargetBrowser TargetBrowser
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTargetBrowser>(this, "TargetBrowser");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TargetBrowser", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpHTMLVersion HTMLVersion
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpHTMLVersion>(this, "HTMLVersion");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "HTMLVersion", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void UseDefaultFolderSuffix()
		{
			 Factory.ExecuteMethod(this, "UseDefaultFolderSuffix");
		}

		#endregion

		#pragma warning restore
	}
}
