using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface Page 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Page : COMObject
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
                    _type = typeof(Page);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Page(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Page(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Page(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Page(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Page(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Page(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Page() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Page(string progId) : base(progId)
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
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shapes Shapes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Shapes>(this, "Shapes", NetOffice.PublisherApi.Shapes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public string PageNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PageNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PageNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbPageType PageType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbPageType>(this, "PageType");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Page Master
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Page>(this, "Master", NetOffice.PublisherApi.Page.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Master", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 PageID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PageID");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 PageIndex
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PageIndex");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.RulerGuides RulerGuides
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.RulerGuides>(this, "RulerGuides", NetOffice.PublisherApi.RulerGuides.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool IgnoreMaster
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IgnoreMaster");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IgnoreMaster", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Wizard Wizard
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Wizard>(this, "Wizard", NetOffice.PublisherApi.Wizard.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Tags Tags
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Tags>(this, "Tags", NetOffice.PublisherApi.Tags.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single XOffsetWithinReaderSpread
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "XOffsetWithinReaderSpread");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single YOffsetWithinReaderSpread
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "YOffsetWithinReaderSpread");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ReaderSpread ReaderSpread
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ReaderSpread>(this, "ReaderSpread", NetOffice.PublisherApi.ReaderSpread.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 Width
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Width");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 Height
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Height");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool IsTwoPageMaster
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsTwoPageMaster");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IsTwoPageMaster", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool IsTrailing
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsTrailing");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool IsLeading
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsLeading");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.HeaderFooter Header
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.HeaderFooter>(this, "Header", NetOffice.PublisherApi.HeaderFooter.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.HeaderFooter Footer
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.HeaderFooter>(this, "Footer", NetOffice.PublisherApi.HeaderFooter.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.PageBackground Background
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.PageBackground>(this, "Background", NetOffice.PublisherApi.PageBackground.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.WebPageOptions WebPageOptions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.WebPageOptions>(this, "WebPageOptions", NetOffice.PublisherApi.WebPageOptions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.LayoutGuides LayoutGuides
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.LayoutGuides>(this, "LayoutGuides", NetOffice.PublisherApi.LayoutGuides.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool IsWizardPage
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsWizardPage");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Publisher", 14,15,16)]
		public void SaveAsPicture10(string filename)
		{
			 Factory.ExecuteMethod(this, "SaveAsPicture10", filename);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="newAbbreviation">optional string NewAbbreviation = </param>
		/// <param name="newName">optional string NewName = </param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Page Duplicate(object newAbbreviation, object newName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Page>(this, "Duplicate", NetOffice.PublisherApi.Page.LateBindingApiWrapperType, newAbbreviation, newName);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Page Duplicate()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Page>(this, "Duplicate", NetOffice.PublisherApi.Page.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="newAbbreviation">optional string NewAbbreviation = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Page Duplicate(object newAbbreviation)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Page>(this, "Duplicate", NetOffice.PublisherApi.Page.LateBindingApiWrapperType, newAbbreviation);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="page">Int32 page</param>
		/// <param name="after">optional bool After = true</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void Move(Int32 page, object after)
		{
			 Factory.ExecuteMethod(this, "Move", page, after);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="page">Int32 page</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void Move(Int32 page)
		{
			 Factory.ExecuteMethod(this, "Move", page);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="pbResolution">optional NetOffice.PublisherApi.Enums.PbPictureResolution pbResolution = 0</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void SaveAsPicture(string filename, object pbResolution)
		{
			 Factory.ExecuteMethod(this, "SaveAsPicture", filename, pbResolution);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void SaveAsPicture(string filename)
		{
			 Factory.ExecuteMethod(this, "SaveAsPicture", filename);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void ExportEmailHTML(string filename)
		{
			 Factory.ExecuteMethod(this, "ExportEmailHTML", filename);
		}

		#endregion

		#pragma warning restore
	}
}
