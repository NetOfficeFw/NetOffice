using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface Options 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Options : COMObject
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
                    _type = typeof(Options);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Options(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Options(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Options(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Options(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Options(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Options(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Options() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Options(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool AllowBackgroundSave
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowBackgroundSave");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowBackgroundSave", value);
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool AutoFormatWord
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatWord");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatWord", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool AutoHyphenate
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoHyphenate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoHyphenate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool AutoSelectWord
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoSelectWord");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoSelectWord", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool DisplayPrintTroubleshooter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayPrintTroubleshooter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayPrintTroubleshooter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool DisplayStatusBar
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayStatusBar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayStatusBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool DragAndDropText
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DragAndDropText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DragAndDropText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.PublisherApi.Enums.PbPlacementType EnvelopePrintPlacement
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbPlacementType>(this, "EnvelopePrintPlacement");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "EnvelopePrintPlacement", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.PublisherApi.Enums.PbOrientationType EnvelopePrintOrientation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbOrientationType>(this, "EnvelopePrintOrientation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "EnvelopePrintOrientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public object HyphenationZone
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "HyphenationZone");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "HyphenationZone", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbUnitType MeasurementUnit
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbUnitType>(this, "MeasurementUnit");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MeasurementUnit", value);
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
		public string PathForPictures
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PathForPictures");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PathForPictures", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public string PathForPublications
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PathForPublications");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PathForPublications", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool PrintLineByLine
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintLineByLine");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintLineByLine", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool SaveAutoRecoverInfo
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SaveAutoRecoverInfo");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SaveAutoRecoverInfo", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 SaveAutoRecoverInfoInterval
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SaveAutoRecoverInfoInterval");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SaveAutoRecoverInfoInterval", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool ShowBasicColors
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowBasicColors");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowBasicColors", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool ShowScreenTipsOnObjects
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowScreenTipsOnObjects");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowScreenTipsOnObjects", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool ShowTipPages
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowTipPages");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowTipPages", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool UpdatePersonalInfoOnSave
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UpdatePersonalInfoOnSave");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UpdatePersonalInfoOnSave", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool UseCatalogAtStartup
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseCatalogAtStartup");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseCatalogAtStartup", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool UseEnvelopePaperSizes
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseEnvelopePaperSizes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseEnvelopePaperSizes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool UseEnvelopePrintOptions
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseEnvelopePrintOptions");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseEnvelopePrintOptions", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool UseHelpfulMousePointers
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseHelpfulMousePointers");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseHelpfulMousePointers", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbDirectionType DefaultPubDirection
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbDirectionType>(this, "DefaultPubDirection");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DefaultPubDirection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool SequenceCheck
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SequenceCheck");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SequenceCheck", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool TypeNReplace
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TypeNReplace");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TypeNReplace", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool AutoKeyboardSwitching
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoKeyboardSwitching");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoKeyboardSwitching", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbDirectionType DefaultTextFlowDirection
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbDirectionType>(this, "DefaultTextFlowDirection");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DefaultTextFlowDirection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool AddHebDoubleQuote
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AddHebDoubleQuote");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AddHebDoubleQuote", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool UseWizardForBlankPublication
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseWizardForBlankPublication");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseWizardForBlankPublication", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void ResetTips()
		{
			 Factory.ExecuteMethod(this, "ResetTips");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void ResetWizardSynchronizing()
		{
			 Factory.ExecuteMethod(this, "ResetWizardSynchronizing");
		}

		#endregion

		#pragma warning restore
	}
}
