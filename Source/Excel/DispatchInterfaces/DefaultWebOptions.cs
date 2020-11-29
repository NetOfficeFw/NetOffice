using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface DefaultWebOptions 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.DefaultWebOptions"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class DefaultWebOptions : COMObject
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
                    _type = typeof(DefaultWebOptions);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public DefaultWebOptions(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public DefaultWebOptions(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DefaultWebOptions(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DefaultWebOptions(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DefaultWebOptions(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DefaultWebOptions(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DefaultWebOptions() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DefaultWebOptions(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.DefaultWebOptions.Application"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.DefaultWebOptions.Creator"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.DefaultWebOptions.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.DefaultWebOptions.RelyOnCSS"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool RelyOnCSS
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RelyOnCSS");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RelyOnCSS", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.DefaultWebOptions.SaveHiddenData"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool SaveHiddenData
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SaveHiddenData");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SaveHiddenData", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.DefaultWebOptions.LoadPictures"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool LoadPictures
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "LoadPictures");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LoadPictures", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.DefaultWebOptions.OrganizeInFolder"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool OrganizeInFolder
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "OrganizeInFolder");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OrganizeInFolder", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.DefaultWebOptions.UpdateLinksOnSave"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool UpdateLinksOnSave
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UpdateLinksOnSave");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UpdateLinksOnSave", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.DefaultWebOptions.UseLongFileNames"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool UseLongFileNames
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseLongFileNames");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseLongFileNames", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.DefaultWebOptions.CheckIfOfficeIsHTMLEditor"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool CheckIfOfficeIsHTMLEditor
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CheckIfOfficeIsHTMLEditor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CheckIfOfficeIsHTMLEditor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.DefaultWebOptions.DownloadComponents"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool DownloadComponents
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DownloadComponents");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DownloadComponents", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.DefaultWebOptions.RelyOnVML"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool RelyOnVML
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RelyOnVML");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RelyOnVML", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.DefaultWebOptions.AllowPNG"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool AllowPNG
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowPNG");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowPNG", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.DefaultWebOptions.ScreenSize"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.DefaultWebOptions.PixelsPerInch"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PixelsPerInch
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PixelsPerInch");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PixelsPerInch", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.DefaultWebOptions.LocationOfComponents"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string LocationOfComponents
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "LocationOfComponents");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LocationOfComponents", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.DefaultWebOptions.Encoding"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.DefaultWebOptions.AlwaysSaveInDefaultEncoding"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool AlwaysSaveInDefaultEncoding
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AlwaysSaveInDefaultEncoding");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AlwaysSaveInDefaultEncoding", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.DefaultWebOptions.Fonts"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.WebPageFonts Fonts
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.WebPageFonts>(this, "Fonts", NetOffice.OfficeApi.WebPageFonts.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.DefaultWebOptions.FolderSuffix"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string FolderSuffix
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FolderSuffix");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.DefaultWebOptions.TargetBrowser"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.DefaultWebOptions.SaveNewWebPagesAsWebArchives"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool SaveNewWebPagesAsWebArchives
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SaveNewWebPagesAsWebArchives");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SaveNewWebPagesAsWebArchives", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
