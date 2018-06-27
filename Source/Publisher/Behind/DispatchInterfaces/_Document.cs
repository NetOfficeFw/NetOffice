using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.PublisherApi;

namespace NetOffice.PublisherApi.Behind
{
	/// <summary>
	/// DispatchInterface _Document 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Document : COMObject, NetOffice.PublisherApi._Document
	{
		#pragma warning disable

		#region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.PublisherApi._Document);
                return _contractType;
            }
        }
        private static Type _contractType;


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

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _Document() : base()
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
		public virtual string ActivePrinter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ActivePrinter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ActivePrinter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Window ActiveWindow
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Window>(this, "ActiveWindow", typeof(NetOffice.PublisherApi.Window));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Application>(this, "Application", typeof(NetOffice.PublisherApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Enums.PbColorMode ColorMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbColorMode>(this, "ColorMode");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.ColorScheme ColorScheme
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ColorScheme>(this, "ColorScheme", typeof(NetOffice.PublisherApi.ColorScheme));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "ColorScheme", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual object DefaultTabStop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "DefaultTabStop");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "DefaultTabStop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool EnvelopeVisible
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EnvelopeVisible");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnvelopeVisible", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual string FullName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FullName");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.LayoutGuides LayoutGuides
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.LayoutGuides>(this, "LayoutGuides", typeof(NetOffice.PublisherApi.LayoutGuides));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.MsoEnvelope MailEnvelope
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.MsoEnvelope>(this, "MailEnvelope", typeof(NetOffice.OfficeApi.MsoEnvelope));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.MailMerge MailMerge
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.MailMerge>(this, "MailMerge", typeof(NetOffice.PublisherApi.MailMerge));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.MasterPages MasterPages
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.MasterPages>(this, "MasterPages", typeof(NetOffice.PublisherApi.MasterPages));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual string Name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Pages Pages
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Pages>(this, "Pages", typeof(NetOffice.PublisherApi.Pages));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.PageSetup PageSetup
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.PageSetup>(this, "PageSetup", typeof(NetOffice.PublisherApi.PageSetup));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual string Path
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Path");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.PublisherApi.Enums.PbPersonalInfoSet PersonalInformationSet
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbPersonalInfoSet>(this, "PersonalInformationSet");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "PersonalInformationSet", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Plates Plates
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Plates>(this, "Plates", typeof(NetOffice.PublisherApi.Plates));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool ReadOnly
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ReadOnly");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Enums.PbDirectionType DocumentDirection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbDirectionType>(this, "DocumentDirection");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "DocumentDirection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool Saved
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Saved");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Enums.PbFileFormat SaveFormat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbFileFormat>(this, "SaveFormat");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.ScratchArea ScratchArea
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ScratchArea>(this, "ScratchArea", typeof(NetOffice.PublisherApi.ScratchArea));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Selection Selection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Selection>(this, "Selection", typeof(NetOffice.PublisherApi.Selection));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Stories Stories
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Stories>(this, "Stories", typeof(NetOffice.PublisherApi.Stories));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Tags Tags
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Tags>(this, "Tags", typeof(NetOffice.PublisherApi.Tags));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.TextStyles TextStyles
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.TextStyles>(this, "TextStyles", typeof(NetOffice.PublisherApi.TextStyles));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool ViewBoundariesAndGuides
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ViewBoundariesAndGuides");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ViewBoundariesAndGuides", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool ViewTwoPageSpread
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ViewTwoPageSpread");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ViewTwoPageSpread", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Wizard Wizard
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Wizard>(this, "Wizard", typeof(NetOffice.PublisherApi.Wizard));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.View ActiveView
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.View>(this, "ActiveView", typeof(NetOffice.PublisherApi.View));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.AdvancedPrintOptions AdvancedPrintOptions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.AdvancedPrintOptions>(this, "AdvancedPrintOptions", typeof(NetOffice.PublisherApi.AdvancedPrintOptions));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.BorderArts BorderArts
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.BorderArts>(this, "BorderArts", typeof(NetOffice.PublisherApi.BorderArts));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool IsDataSourceConnected
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsDataSourceConnected");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.FindReplace Find
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.FindReplace>(this, "Find", typeof(NetOffice.PublisherApi.FindReplace));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Int32 UndoActionsAvailable
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "UndoActionsAvailable");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Int32 RedoActionsAvailable
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "RedoActionsAvailable");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool ViewHorizontalBaseLineGuides
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ViewHorizontalBaseLineGuides");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ViewHorizontalBaseLineGuides", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool ViewVerticalBaseLineGuides
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ViewVerticalBaseLineGuides");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ViewVerticalBaseLineGuides", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Enums.PbPublicationType PublicationType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbPublicationType>(this, "PublicationType");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Sections Sections
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Sections>(this, "Sections", typeof(NetOffice.PublisherApi.Sections));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.WebNavigationBarSets WebNavigationBarSets
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.WebNavigationBarSets>(this, "WebNavigationBarSets", typeof(NetOffice.PublisherApi.WebNavigationBarSets));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool RemovePersonalInformation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "RemovePersonalInformation");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RemovePersonalInformation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool PrintPageBackgrounds
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PrintPageBackgrounds");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PrintPageBackgrounds", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.ColorsInUse ColorsInUse
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ColorsInUse>(this, "ColorsInUse", typeof(NetOffice.PublisherApi.ColorsInUse));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool IsWizard
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsWizard");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.ShapeRange SurplusShapes
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ShapeRange>(this, "SurplusShapes", typeof(NetOffice.PublisherApi.ShapeRange));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Enums.PbPrintStyle PrintStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbPrintStyle>(this, "PrintStyle");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool ViewBoundaries
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ViewBoundaries");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ViewBoundaries", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool ViewGuides
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ViewGuides");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ViewGuides", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.BuildingBlocks AvailableBuildingBlocks
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.BuildingBlocks>(this, "AvailableBuildingBlocks", typeof(NetOffice.PublisherApi.BuildingBlocks));
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void Close()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="mode">NetOffice.PublisherApi.Enums.PbColorMode mode</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Plates CreatePlateCollection(NetOffice.PublisherApi.Enums.PbColorMode mode)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Plates>(this, "CreatePlateCollection", typeof(NetOffice.PublisherApi.Plates), mode);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="mode">NetOffice.PublisherApi.Enums.PbColorMode mode</param>
		/// <param name="plates">optional object plates</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void EnterColorMode10(NetOffice.PublisherApi.Enums.PbColorMode mode, object plates)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EnterColorMode10", mode, plates);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="mode">NetOffice.PublisherApi.Enums.PbColorMode mode</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void EnterColorMode10(NetOffice.PublisherApi.Enums.PbColorMode mode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EnterColorMode10", mode);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="tagName">string tagName</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.ShapeRange FindShapesByTag(string tagName)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.ShapeRange>(this, "FindShapesByTag", typeof(NetOffice.PublisherApi.ShapeRange), tagName);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizardTag">NetOffice.PublisherApi.Enums.PbWizardTag wizardTag</param>
		/// <param name="instance">optional Int32 Instance = -1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.ShapeRange FindShapeByWizardTag(NetOffice.PublisherApi.Enums.PbWizardTag wizardTag, object instance)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.ShapeRange>(this, "FindShapeByWizardTag", typeof(NetOffice.PublisherApi.ShapeRange), wizardTag, instance);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizardTag">NetOffice.PublisherApi.Enums.PbWizardTag wizardTag</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.ShapeRange FindShapeByWizardTag(NetOffice.PublisherApi.Enums.PbWizardTag wizardTag)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.ShapeRange>(this, "FindShapeByWizardTag", typeof(NetOffice.PublisherApi.ShapeRange), wizardTag);
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
		public virtual void PrintOut(object from, object to, object printToFile, object copies, object collate)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ from, to, printToFile, copies, collate });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void PrintOut()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void PrintOut(object from)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", from);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void PrintOut(object from, object to)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", from, to);
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
		public virtual void PrintOut(object from, object to, object printToFile)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", from, to, printToFile);
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
		public virtual void PrintOut(object from, object to, object printToFile, object copies)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", from, to, printToFile, copies);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void Save()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Save");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">optional object filename</param>
		/// <param name="format">optional NetOffice.PublisherApi.Enums.PbFileFormat Format = 1</param>
		/// <param name="addToRecentFiles">optional bool AddToRecentFiles = true</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void SaveAs(object filename, object format, object addToRecentFiles)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", filename, format, addToRecentFiles);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void SaveAs()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">optional object filename</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void SaveAs(object filename)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", filename);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">optional object filename</param>
		/// <param name="format">optional NetOffice.PublisherApi.Enums.PbFileFormat Format = 1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void SaveAs(object filename, object format)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", filename, format);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="oh">Int32 oh</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void SelectID(Int32 oh)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SelectID", oh);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void UndoClear()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "UndoClear");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void UpdateOLEObjects()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateOLEObjects");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="count">optional Int32 Count = 1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void Undo(object count)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Undo", count);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void Undo()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Undo");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="count">optional Int32 Count = 1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void Redo(object count)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Redo", count);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void Redo()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Redo");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="actionName">string actionName</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void BeginCustomUndoAction(string actionName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "BeginCustomUndoAction", actionName);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void EndCustomUndoAction()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EndCustomUndoAction");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void WebPagePreview()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "WebPagePreview");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">NetOffice.PublisherApi.Enums.PbPublicationType value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ConvertPublicationType(NetOffice.PublisherApi.Enums.PbPublicationType value)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ConvertPublicationType", value);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="mode">NetOffice.PublisherApi.Enums.PbColorMode mode</param>
		/// <param name="plates">optional object plates</param>
		/// <param name="deleteExcessInks">optional bool DeleteExcessInks = false</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void EnterColorMode(NetOffice.PublisherApi.Enums.PbColorMode mode, object plates, object deleteExcessInks)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EnterColorMode", mode, plates, deleteExcessInks);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="mode">NetOffice.PublisherApi.Enums.PbColorMode mode</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void EnterColorMode(NetOffice.PublisherApi.Enums.PbColorMode mode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EnterColorMode", mode);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="mode">NetOffice.PublisherApi.Enums.PbColorMode mode</param>
		/// <param name="plates">optional object plates</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void EnterColorMode(NetOffice.PublisherApi.Enums.PbColorMode mode, object plates)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EnterColorMode", mode, plates);
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
		public virtual void PrintOutEx(object from, object to, object printToFile, object copies, object collate, object printStyle)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutEx", new object[]{ from, to, printToFile, copies, collate, printStyle });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void PrintOutEx()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutEx");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void PrintOutEx(object from)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutEx", from);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void PrintOutEx(object from, object to)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutEx", from, to);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void PrintOutEx(object from, object to, object printToFile)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutEx", from, to, printToFile);
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
		public virtual void PrintOutEx(object from, object to, object printToFile, object copies)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutEx", from, to, printToFile, copies);
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
		public virtual void PrintOutEx(object from, object to, object printToFile, object copies, object collate)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutEx", new object[]{ from, to, printToFile, copies, collate });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizard">NetOffice.PublisherApi.Enums.PbWizard wizard</param>
		/// <param name="design">optional Int32 Design = -1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ChangeDocument(NetOffice.PublisherApi.Enums.PbWizard wizard, object design)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ChangeDocument", wizard, design);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizard">NetOffice.PublisherApi.Enums.PbWizard wizard</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ChangeDocument(NetOffice.PublisherApi.Enums.PbWizard wizard)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ChangeDocument", wizard);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void SetBusinessInformation(string name)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetBusinessInformation", name);
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
		public virtual void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies, object collate, object printStyle, object docStructureTags, object bitmapMissingFonts, object useISO19005_1, object externalExporter)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold, from, to, copies, collate, printStyle, docStructureTags, bitmapMissingFonts, useISO19005_1, externalExporter });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", format, filename);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", format, filename, intent);
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
		public virtual void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", format, filename, intent, includeDocumentProperties);
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
		public virtual void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ format, filename, intent, includeDocumentProperties, colorDownsampleTarget });
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
		public virtual void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold });
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
		public virtual void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget });
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
		public virtual void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold });
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
		public virtual void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold, from });
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
		public virtual void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold, from, to });
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
		public virtual void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold, from, to, copies });
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
		public virtual void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies, object collate)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold, from, to, copies, collate });
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
		public virtual void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies, object collate, object printStyle)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold, from, to, copies, collate, printStyle });
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
		public virtual void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies, object collate, object printStyle, object docStructureTags)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold, from, to, copies, collate, printStyle, docStructureTags });
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
		public virtual void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies, object collate, object printStyle, object docStructureTags, object bitmapMissingFonts)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold, from, to, copies, collate, printStyle, docStructureTags, bitmapMissingFonts });
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
		public virtual void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies, object collate, object printStyle, object docStructureTags, object bitmapMissingFonts, object useISO19005_1)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold, from, to, copies, collate, printStyle, docStructureTags, bitmapMissingFonts, useISO19005_1 });
		}

		#endregion

		#pragma warning restore
	}
}


