using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface AdvancedPrintOptions 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class AdvancedPrintOptions : COMObject
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
                    _type = typeof(AdvancedPrintOptions);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public AdvancedPrintOptions(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public AdvancedPrintOptions(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AdvancedPrintOptions(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AdvancedPrintOptions(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AdvancedPrintOptions(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AdvancedPrintOptions(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AdvancedPrintOptions() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AdvancedPrintOptions(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

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
		public bool VerticalFlip
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "VerticalFlip");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "VerticalFlip", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool HorizontalFlip
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HorizontalFlip");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HorizontalFlip", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool NegativeImage
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "NegativeImage");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NegativeImage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool UseOnlyPublicationFonts
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseOnlyPublicationFonts");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseOnlyPublicationFonts", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool PrintCropMarks
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintCropMarks");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintCropMarks", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool PrintRegistrationMarks
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintRegistrationMarks");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintRegistrationMarks", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool PrintJobInformation
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintJobInformation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintJobInformation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool PrintDensityBars
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintDensityBars");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintDensityBars", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool PrintColorBars
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintColorBars");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintColorBars", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool AllowBleeds
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowBleeds");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowBleeds", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool PrintBleedMarks
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintBleedMarks");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintBleedMarks", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool PrintBlankPlates
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintBlankPlates");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintBlankPlates", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbPrintGraphics GraphicsResolution
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbPrintGraphics>(this, "GraphicsResolution");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "GraphicsResolution", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public string Resolution
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Resolution");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Resolution", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.PublisherApi.PrintableRect PrintableRect
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.PrintableRect>(this, "PrintableRect", NetOffice.PublisherApi.PrintableRect.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbInksToPrint InksToPrint
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbInksToPrint>(this, "InksToPrint");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "InksToPrint", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.PublisherApi.Enums.PbPrintMode PrintMode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbPrintMode>(this, "PrintMode");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PrintMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.PrintablePlates PrintablePlates
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.PrintablePlates>(this, "PrintablePlates", NetOffice.PublisherApi.PrintablePlates.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool IsPostscriptPrinter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsPostscriptPrinter");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool UseCustomHalftone
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseCustomHalftone");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseCustomHalftone", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool PrintCMYKByDefault
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintCMYKByDefault");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintCMYKByDefault", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool BackSideInsertFaceUp
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "BackSideInsertFaceUp");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BackSideInsertFaceUp", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbPlacementType ManualFeedAlign
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbPlacementType>(this, "ManualFeedAlign");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ManualFeedAlign", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbOrientationType ManualFeedDirection
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbOrientationType>(this, "ManualFeedDirection");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ManualFeedDirection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool PageRotated
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PageRotated");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PageRotated", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
