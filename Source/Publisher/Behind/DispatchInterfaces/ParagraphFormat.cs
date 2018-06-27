using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.PublisherApi;

namespace NetOffice.PublisherApi.Behind
{
	/// <summary>
	/// DispatchInterface ParagraphFormat 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ParagraphFormat : COMObject, NetOffice.PublisherApi.ParagraphFormat
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
                    _contractType = typeof(NetOffice.PublisherApi.ParagraphFormat);
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
                    _type = typeof(ParagraphFormat);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ParagraphFormat() : base()
		{

		}

		#endregion
		
		#region Properties

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
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Enums.PbParagraphAlignmentType Alignment
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbParagraphAlignmentType>(this, "Alignment");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Alignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual object FirstLineIndent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "FirstLineIndent");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "FirstLineIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual object LeftIndent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "LeftIndent");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "LeftIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual object RightIndent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "RightIndent");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "RightIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual object SpaceAfter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "SpaceAfter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "SpaceAfter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual object SpaceBefore
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "SpaceBefore");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "SpaceBefore", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual object LineSpacing
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "LineSpacing");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "LineSpacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Enums.PbLineSpacingRule LineSpacingRule
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbLineSpacingRule>(this, "LineSpacingRule");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "LineSpacingRule", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.TabStops Tabs
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.TabStops>(this, "Tabs", typeof(NetOffice.PublisherApi.TabStops));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Enums.PbTextDirection TextDirection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbTextDirection>(this, "TextDirection");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "TextDirection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual object TextStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "TextStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "TextStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool AttachedToText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AttachedToText");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Int32 KashidaPercentage
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "KashidaPercentage");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "KashidaPercentage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Enums.PbListType ListType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbListType>(this, "ListType");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Single ListIndent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "ListIndent");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ListIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual string ListBulletText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ListBulletText");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Single ListBulletFontSize
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "ListBulletFontSize");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ListBulletFontSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual string ListBulletFontName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ListBulletFontName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ListBulletFontName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Enums.PbListSeparator ListNumberSeparator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbListSeparator>(this, "ListNumberSeparator");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ListNumberSeparator", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Int32 ListNumberStart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ListNumberStart");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ListNumberStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Int32 CharBasedFirstLineIndent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CharBasedFirstLineIndent");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CharBasedFirstLineIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState UseCharBasedFirstLineIndent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "UseCharBasedFirstLineIndent");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "UseCharBasedFirstLineIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState WidowControl
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "WidowControl");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "WidowControl", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState KeepLinesTogether
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "KeepLinesTogether");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "KeepLinesTogether", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState KeepWithNext
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "KeepWithNext");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "KeepWithNext", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState StartInNextTextBox
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "StartInNextTextBox");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "StartInNextTextBox", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState LockToBaseLine
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "LockToBaseLine");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "LockToBaseLine", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void Reset()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Reset");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.ParagraphFormat Duplicate()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.ParagraphFormat>(this, "Duplicate", typeof(NetOffice.PublisherApi.ParagraphFormat));
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="rule">NetOffice.PublisherApi.Enums.PbLineSpacingRule rule</param>
		/// <param name="spacing">optional object spacing</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void SetLineSpacing(NetOffice.PublisherApi.Enums.PbLineSpacingRule rule, object spacing)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetLineSpacing", rule, spacing);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="rule">NetOffice.PublisherApi.Enums.PbLineSpacingRule rule</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void SetLineSpacing(NetOffice.PublisherApi.Enums.PbLineSpacingRule rule)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetLineSpacing", rule);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">NetOffice.PublisherApi.Enums.PbListType value</param>
		/// <param name="bulletText">optional string BulletText = </param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void SetListType(NetOffice.PublisherApi.Enums.PbListType value, object bulletText)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetListType", value, bulletText);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">NetOffice.PublisherApi.Enums.PbListType value</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void SetListType(NetOffice.PublisherApi.Enums.PbListType value)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetListType", value);
		}

		#endregion

		#pragma warning restore
	}
}


