using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface ParagraphFormat 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ParagraphFormat : COMObject
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
                    _type = typeof(ParagraphFormat);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ParagraphFormat(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ParagraphFormat(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ParagraphFormat(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ParagraphFormat(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ParagraphFormat(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ParagraphFormat(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ParagraphFormat() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ParagraphFormat(string progId) : base(progId)
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
		public NetOffice.PublisherApi.Enums.PbParagraphAlignmentType Alignment
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbParagraphAlignmentType>(this, "Alignment");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Alignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public object FirstLineIndent
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "FirstLineIndent");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "FirstLineIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public object LeftIndent
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "LeftIndent");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "LeftIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public object RightIndent
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "RightIndent");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "RightIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public object SpaceAfter
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "SpaceAfter");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "SpaceAfter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public object SpaceBefore
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "SpaceBefore");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "SpaceBefore", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public object LineSpacing
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "LineSpacing");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "LineSpacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbLineSpacingRule LineSpacingRule
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbLineSpacingRule>(this, "LineSpacingRule");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "LineSpacingRule", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TabStops Tabs
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.TabStops>(this, "Tabs", NetOffice.PublisherApi.TabStops.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbTextDirection TextDirection
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbTextDirection>(this, "TextDirection");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TextDirection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public object TextStyle
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "TextStyle");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "TextStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool AttachedToText
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AttachedToText");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 KashidaPercentage
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "KashidaPercentage");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "KashidaPercentage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbListType ListType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbListType>(this, "ListType");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single ListIndent
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "ListIndent");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ListIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public string ListBulletText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ListBulletText");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single ListBulletFontSize
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "ListBulletFontSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ListBulletFontSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public string ListBulletFontName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ListBulletFontName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ListBulletFontName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbListSeparator ListNumberSeparator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbListSeparator>(this, "ListNumberSeparator");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ListNumberSeparator", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 ListNumberStart
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ListNumberStart");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ListNumberStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 CharBasedFirstLineIndent
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CharBasedFirstLineIndent");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CharBasedFirstLineIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState UseCharBasedFirstLineIndent
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "UseCharBasedFirstLineIndent");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "UseCharBasedFirstLineIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState WidowControl
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "WidowControl");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "WidowControl", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState KeepLinesTogether
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "KeepLinesTogether");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "KeepLinesTogether", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState KeepWithNext
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "KeepWithNext");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "KeepWithNext", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState StartInNextTextBox
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "StartInNextTextBox");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "StartInNextTextBox", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState LockToBaseLine
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "LockToBaseLine");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "LockToBaseLine", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void Reset()
		{
			 Factory.ExecuteMethod(this, "Reset");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ParagraphFormat Duplicate()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.ParagraphFormat>(this, "Duplicate", NetOffice.PublisherApi.ParagraphFormat.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="rule">NetOffice.PublisherApi.Enums.PbLineSpacingRule rule</param>
		/// <param name="spacing">optional object spacing</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void SetLineSpacing(NetOffice.PublisherApi.Enums.PbLineSpacingRule rule, object spacing)
		{
			 Factory.ExecuteMethod(this, "SetLineSpacing", rule, spacing);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="rule">NetOffice.PublisherApi.Enums.PbLineSpacingRule rule</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void SetLineSpacing(NetOffice.PublisherApi.Enums.PbLineSpacingRule rule)
		{
			 Factory.ExecuteMethod(this, "SetLineSpacing", rule);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">NetOffice.PublisherApi.Enums.PbListType value</param>
		/// <param name="bulletText">optional string BulletText = </param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void SetListType(NetOffice.PublisherApi.Enums.PbListType value, object bulletText)
		{
			 Factory.ExecuteMethod(this, "SetListType", value, bulletText);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">NetOffice.PublisherApi.Enums.PbListType value</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void SetListType(NetOffice.PublisherApi.Enums.PbListType value)
		{
			 Factory.ExecuteMethod(this, "SetListType", value);
		}

		#endregion

		#pragma warning restore
	}
}
