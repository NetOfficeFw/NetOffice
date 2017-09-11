using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface _Document 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Document : COMObject
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
                    _type = typeof(_Document);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _Document(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Document(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Document(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Document(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Document(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Document(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Document() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Document(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string ActivePrinter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ActivePrinter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ActivePrinter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Window ActiveWindow
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Window>(this, "ActiveWindow", NetOffice.PublisherApi.Window.LateBindingApiWrapperType);
			}
		}

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
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbColorMode ColorMode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbColorMode>(this, "ColorMode");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ColorScheme ColorScheme
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ColorScheme>(this, "ColorScheme", NetOffice.PublisherApi.ColorScheme.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "ColorScheme", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public object DefaultTabStop
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DefaultTabStop");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "DefaultTabStop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool EnvelopeVisible
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnvelopeVisible");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnvelopeVisible", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public string FullName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FullName");
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
		public NetOffice.OfficeApi.MsoEnvelope MailEnvelope
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.MsoEnvelope>(this, "MailEnvelope", NetOffice.OfficeApi.MsoEnvelope.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.MailMerge MailMerge
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.MailMerge>(this, "MailMerge", NetOffice.PublisherApi.MailMerge.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.MasterPages MasterPages
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.MasterPages>(this, "MasterPages", NetOffice.PublisherApi.MasterPages.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Pages Pages
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Pages>(this, "Pages", NetOffice.PublisherApi.Pages.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.PageSetup PageSetup
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.PageSetup>(this, "PageSetup", NetOffice.PublisherApi.PageSetup.LateBindingApiWrapperType);
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
		public string Path
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Path");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.PublisherApi.Enums.PbPersonalInfoSet PersonalInformationSet
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbPersonalInfoSet>(this, "PersonalInformationSet");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PersonalInformationSet", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Plates Plates
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Plates>(this, "Plates", NetOffice.PublisherApi.Plates.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool ReadOnly
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ReadOnly");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbDirectionType DocumentDirection
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbDirectionType>(this, "DocumentDirection");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DocumentDirection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool Saved
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Saved");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbFileFormat SaveFormat
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbFileFormat>(this, "SaveFormat");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ScratchArea ScratchArea
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ScratchArea>(this, "ScratchArea", NetOffice.PublisherApi.ScratchArea.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Selection Selection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Selection>(this, "Selection", NetOffice.PublisherApi.Selection.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Stories Stories
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Stories>(this, "Stories", NetOffice.PublisherApi.Stories.LateBindingApiWrapperType);
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
		public NetOffice.PublisherApi.TextStyles TextStyles
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.TextStyles>(this, "TextStyles", NetOffice.PublisherApi.TextStyles.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool ViewBoundariesAndGuides
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ViewBoundariesAndGuides");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ViewBoundariesAndGuides", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool ViewTwoPageSpread
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ViewTwoPageSpread");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ViewTwoPageSpread", value);
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
		public NetOffice.PublisherApi.View ActiveView
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.View>(this, "ActiveView", NetOffice.PublisherApi.View.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.AdvancedPrintOptions AdvancedPrintOptions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.AdvancedPrintOptions>(this, "AdvancedPrintOptions", NetOffice.PublisherApi.AdvancedPrintOptions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.BorderArts BorderArts
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.BorderArts>(this, "BorderArts", NetOffice.PublisherApi.BorderArts.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool IsDataSourceConnected
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsDataSourceConnected");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.FindReplace Find
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.FindReplace>(this, "Find", NetOffice.PublisherApi.FindReplace.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 UndoActionsAvailable
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "UndoActionsAvailable");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 RedoActionsAvailable
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "RedoActionsAvailable");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool ViewHorizontalBaseLineGuides
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ViewHorizontalBaseLineGuides");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ViewHorizontalBaseLineGuides", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool ViewVerticalBaseLineGuides
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ViewVerticalBaseLineGuides");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ViewVerticalBaseLineGuides", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbPublicationType PublicationType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbPublicationType>(this, "PublicationType");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Sections Sections
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Sections>(this, "Sections", NetOffice.PublisherApi.Sections.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.WebNavigationBarSets WebNavigationBarSets
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.WebNavigationBarSets>(this, "WebNavigationBarSets", NetOffice.PublisherApi.WebNavigationBarSets.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool RemovePersonalInformation
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RemovePersonalInformation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RemovePersonalInformation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool PrintPageBackgrounds
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintPageBackgrounds");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintPageBackgrounds", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ColorsInUse ColorsInUse
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ColorsInUse>(this, "ColorsInUse", NetOffice.PublisherApi.ColorsInUse.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool IsWizard
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsWizard");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ShapeRange SurplusShapes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ShapeRange>(this, "SurplusShapes", NetOffice.PublisherApi.ShapeRange.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbPrintStyle PrintStyle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbPrintStyle>(this, "PrintStyle");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool ViewBoundaries
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ViewBoundaries");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ViewBoundaries", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool ViewGuides
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ViewGuides");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ViewGuides", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.BuildingBlocks AvailableBuildingBlocks
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.BuildingBlocks>(this, "AvailableBuildingBlocks", NetOffice.PublisherApi.BuildingBlocks.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void Close()
		{
			 Factory.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="mode">NetOffice.PublisherApi.Enums.PbColorMode mode</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Plates CreatePlateCollection(NetOffice.PublisherApi.Enums.PbColorMode mode)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Plates>(this, "CreatePlateCollection", NetOffice.PublisherApi.Plates.LateBindingApiWrapperType, mode);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="mode">NetOffice.PublisherApi.Enums.PbColorMode mode</param>
		/// <param name="plates">optional object plates</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Publisher", 14,15,16)]
		public void EnterColorMode10(NetOffice.PublisherApi.Enums.PbColorMode mode, object plates)
		{
			 Factory.ExecuteMethod(this, "EnterColorMode10", mode, plates);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="mode">NetOffice.PublisherApi.Enums.PbColorMode mode</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void EnterColorMode10(NetOffice.PublisherApi.Enums.PbColorMode mode)
		{
			 Factory.ExecuteMethod(this, "EnterColorMode10", mode);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="tagName">string tagName</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ShapeRange FindShapesByTag(string tagName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.ShapeRange>(this, "FindShapesByTag", NetOffice.PublisherApi.ShapeRange.LateBindingApiWrapperType, tagName);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizardTag">NetOffice.PublisherApi.Enums.PbWizardTag wizardTag</param>
		/// <param name="instance">optional Int32 Instance = -1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ShapeRange FindShapeByWizardTag(NetOffice.PublisherApi.Enums.PbWizardTag wizardTag, object instance)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.ShapeRange>(this, "FindShapeByWizardTag", NetOffice.PublisherApi.ShapeRange.LateBindingApiWrapperType, wizardTag, instance);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizardTag">NetOffice.PublisherApi.Enums.PbWizardTag wizardTag</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ShapeRange FindShapeByWizardTag(NetOffice.PublisherApi.Enums.PbWizardTag wizardTag)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.ShapeRange>(this, "FindShapeByWizardTag", NetOffice.PublisherApi.ShapeRange.LateBindingApiWrapperType, wizardTag);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		/// <param name="collate">optional bool Collate = true</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Publisher", 14,15,16)]
		public void PrintOut(object from, object to, object printToFile, object copies, object collate)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ from, to, printToFile, copies, collate });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void PrintOut()
		{
			 Factory.ExecuteMethod(this, "PrintOut");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void PrintOut(object from)
		{
			 Factory.ExecuteMethod(this, "PrintOut", from);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void PrintOut(object from, object to)
		{
			 Factory.ExecuteMethod(this, "PrintOut", from, to);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void PrintOut(object from, object to, object printToFile)
		{
			 Factory.ExecuteMethod(this, "PrintOut", from, to, printToFile);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void PrintOut(object from, object to, object printToFile, object copies)
		{
			 Factory.ExecuteMethod(this, "PrintOut", from, to, printToFile, copies);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void Save()
		{
			 Factory.ExecuteMethod(this, "Save");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">optional object filename</param>
		/// <param name="format">optional NetOffice.PublisherApi.Enums.PbFileFormat Format = 1</param>
		/// <param name="addToRecentFiles">optional bool AddToRecentFiles = true</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void SaveAs(object filename, object format, object addToRecentFiles)
		{
			 Factory.ExecuteMethod(this, "SaveAs", filename, format, addToRecentFiles);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void SaveAs()
		{
			 Factory.ExecuteMethod(this, "SaveAs");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">optional object filename</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void SaveAs(object filename)
		{
			 Factory.ExecuteMethod(this, "SaveAs", filename);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">optional object filename</param>
		/// <param name="format">optional NetOffice.PublisherApi.Enums.PbFileFormat Format = 1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void SaveAs(object filename, object format)
		{
			 Factory.ExecuteMethod(this, "SaveAs", filename, format);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="oh">Int32 oh</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Publisher", 14,15,16)]
		public void SelectID(Int32 oh)
		{
			 Factory.ExecuteMethod(this, "SelectID", oh);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void UndoClear()
		{
			 Factory.ExecuteMethod(this, "UndoClear");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void UpdateOLEObjects()
		{
			 Factory.ExecuteMethod(this, "UpdateOLEObjects");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="count">optional Int32 Count = 1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void Undo(object count)
		{
			 Factory.ExecuteMethod(this, "Undo", count);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void Undo()
		{
			 Factory.ExecuteMethod(this, "Undo");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="count">optional Int32 Count = 1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void Redo(object count)
		{
			 Factory.ExecuteMethod(this, "Redo", count);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void Redo()
		{
			 Factory.ExecuteMethod(this, "Redo");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="actionName">string actionName</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void BeginCustomUndoAction(string actionName)
		{
			 Factory.ExecuteMethod(this, "BeginCustomUndoAction", actionName);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void EndCustomUndoAction()
		{
			 Factory.ExecuteMethod(this, "EndCustomUndoAction");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void WebPagePreview()
		{
			 Factory.ExecuteMethod(this, "WebPagePreview");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">NetOffice.PublisherApi.Enums.PbPublicationType value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void ConvertPublicationType(NetOffice.PublisherApi.Enums.PbPublicationType value)
		{
			 Factory.ExecuteMethod(this, "ConvertPublicationType", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="mode">NetOffice.PublisherApi.Enums.PbColorMode mode</param>
		/// <param name="plates">optional object plates</param>
		/// <param name="deleteExcessInks">optional bool DeleteExcessInks = false</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void EnterColorMode(NetOffice.PublisherApi.Enums.PbColorMode mode, object plates, object deleteExcessInks)
		{
			 Factory.ExecuteMethod(this, "EnterColorMode", mode, plates, deleteExcessInks);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="mode">NetOffice.PublisherApi.Enums.PbColorMode mode</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void EnterColorMode(NetOffice.PublisherApi.Enums.PbColorMode mode)
		{
			 Factory.ExecuteMethod(this, "EnterColorMode", mode);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="mode">NetOffice.PublisherApi.Enums.PbColorMode mode</param>
		/// <param name="plates">optional object plates</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void EnterColorMode(NetOffice.PublisherApi.Enums.PbColorMode mode, object plates)
		{
			 Factory.ExecuteMethod(this, "EnterColorMode", mode, plates);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		/// <param name="collate">optional bool Collate = true</param>
		/// <param name="printStyle">optional NetOffice.PublisherApi.Enums.PbPrintStyle PrintStyle = 0</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void PrintOutEx(object from, object to, object printToFile, object copies, object collate, object printStyle)
		{
			 Factory.ExecuteMethod(this, "PrintOutEx", new object[]{ from, to, printToFile, copies, collate, printStyle });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void PrintOutEx()
		{
			 Factory.ExecuteMethod(this, "PrintOutEx");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void PrintOutEx(object from)
		{
			 Factory.ExecuteMethod(this, "PrintOutEx", from);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void PrintOutEx(object from, object to)
		{
			 Factory.ExecuteMethod(this, "PrintOutEx", from, to);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void PrintOutEx(object from, object to, object printToFile)
		{
			 Factory.ExecuteMethod(this, "PrintOutEx", from, to, printToFile);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void PrintOutEx(object from, object to, object printToFile, object copies)
		{
			 Factory.ExecuteMethod(this, "PrintOutEx", from, to, printToFile, copies);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		/// <param name="collate">optional bool Collate = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void PrintOutEx(object from, object to, object printToFile, object copies, object collate)
		{
			 Factory.ExecuteMethod(this, "PrintOutEx", new object[]{ from, to, printToFile, copies, collate });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizard">NetOffice.PublisherApi.Enums.PbWizard wizard</param>
		/// <param name="design">optional Int32 Design = -1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void ChangeDocument(NetOffice.PublisherApi.Enums.PbWizard wizard, object design)
		{
			 Factory.ExecuteMethod(this, "ChangeDocument", wizard, design);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizard">NetOffice.PublisherApi.Enums.PbWizard wizard</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ChangeDocument(NetOffice.PublisherApi.Enums.PbWizard wizard)
		{
			 Factory.ExecuteMethod(this, "ChangeDocument", wizard);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void SetBusinessInformation(string name)
		{
			 Factory.ExecuteMethod(this, "SetBusinessInformation", name);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		/// <param name="collate">optional bool Collate = true</param>
		/// <param name="printStyle">optional NetOffice.PublisherApi.Enums.PbPrintStyle PrintStyle = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		/// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
		/// <param name="externalExporter">optional object externalExporter</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies, object collate, object printStyle, object docStructureTags, object bitmapMissingFonts, object useISO19005_1, object externalExporter)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold, from, to, copies, collate, printStyle, docStructureTags, bitmapMissingFonts, useISO19005_1, externalExporter });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", format, filename);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", format, filename, intent);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", format, filename, intent, includeDocumentProperties);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ format, filename, intent, includeDocumentProperties, colorDownsampleTarget });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		/// <param name="from">optional Int32 From = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold, from });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold, from, to });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold, from, to, copies });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		/// <param name="collate">optional bool Collate = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies, object collate)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold, from, to, copies, collate });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		/// <param name="collate">optional bool Collate = true</param>
		/// <param name="printStyle">optional NetOffice.PublisherApi.Enums.PbPrintStyle PrintStyle = 0</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies, object collate, object printStyle)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold, from, to, copies, collate, printStyle });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		/// <param name="collate">optional bool Collate = true</param>
		/// <param name="printStyle">optional NetOffice.PublisherApi.Enums.PbPrintStyle PrintStyle = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies, object collate, object printStyle, object docStructureTags)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold, from, to, copies, collate, printStyle, docStructureTags });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		/// <param name="collate">optional bool Collate = true</param>
		/// <param name="printStyle">optional NetOffice.PublisherApi.Enums.PbPrintStyle PrintStyle = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies, object collate, object printStyle, object docStructureTags, object bitmapMissingFonts)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold, from, to, copies, collate, printStyle, docStructureTags, bitmapMissingFonts });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		/// <param name="collate">optional bool Collate = true</param>
		/// <param name="printStyle">optional NetOffice.PublisherApi.Enums.PbPrintStyle PrintStyle = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		/// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies, object collate, object printStyle, object docStructureTags, object bitmapMissingFonts, object useISO19005_1)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold, from, to, copies, collate, printStyle, docStructureTags, bitmapMissingFonts, useISO19005_1 });
		}

		#endregion

		#pragma warning restore
	}
}
