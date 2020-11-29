﻿using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface _LetterContent 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _LetterContent : COMObject
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
                    _type = typeof(_LetterContent);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _LetterContent(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _LetterContent(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _LetterContent(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _LetterContent(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _LetterContent(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _LetterContent(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _LetterContent() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _LetterContent(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.Application"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", NetOffice.WordApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.Creator"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.Parent"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.Duplicate"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.LetterContent Duplicate
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.LetterContent>(this, "Duplicate", NetOffice.WordApi.LetterContent.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.DateFormat"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string DateFormat
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DateFormat");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DateFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.IncludeHeaderFooter"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool IncludeHeaderFooter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IncludeHeaderFooter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IncludeHeaderFooter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.PageDesign"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string PageDesign
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PageDesign");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PageDesign", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.LetterStyle"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdLetterStyle LetterStyle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdLetterStyle>(this, "LetterStyle");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "LetterStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.Letterhead"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Letterhead
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Letterhead");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Letterhead", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.LetterheadLocation"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdLetterheadLocation LetterheadLocation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdLetterheadLocation>(this, "LetterheadLocation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "LetterheadLocation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.LetterheadSize"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single LetterheadSize
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "LetterheadSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LetterheadSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.RecipientName"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string RecipientName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RecipientName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecipientName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.RecipientAddress"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string RecipientAddress
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RecipientAddress");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecipientAddress", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.Salutation"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string Salutation
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Salutation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Salutation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.SalutationType"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdSalutationType SalutationType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdSalutationType>(this, "SalutationType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SalutationType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.RecipientReference"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string RecipientReference
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RecipientReference");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecipientReference", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.MailingInstructions"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string MailingInstructions
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MailingInstructions");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MailingInstructions", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.AttentionLine"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string AttentionLine
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AttentionLine");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AttentionLine", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.Subject"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string Subject
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Subject");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Subject", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.EnclosureNumber"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 EnclosureNumber
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "EnclosureNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnclosureNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.CCList"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string CCList
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CCList");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CCList", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.ReturnAddress"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string ReturnAddress
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ReturnAddress");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReturnAddress", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.SenderName"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string SenderName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SenderName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SenderName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.Closing"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string Closing
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Closing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Closing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.SenderCompany"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string SenderCompany
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SenderCompany");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SenderCompany", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.SenderJobTitle"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string SenderJobTitle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SenderJobTitle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SenderJobTitle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.SenderInitials"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string SenderInitials
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SenderInitials");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SenderInitials", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.InfoBlock"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool InfoBlock
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "InfoBlock");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InfoBlock", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.RecipientCode"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string RecipientCode
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RecipientCode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecipientCode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.RecipientGender"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdSalutationGender RecipientGender
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdSalutationGender>(this, "RecipientGender");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RecipientGender", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.ReturnAddressShortForm"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string ReturnAddressShortForm
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ReturnAddressShortForm");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReturnAddressShortForm", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.SenderCity"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string SenderCity
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SenderCity");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SenderCity", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.SenderCode"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string SenderCode
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SenderCode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SenderCode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.SenderGender"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdSalutationGender SenderGender
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdSalutationGender>(this, "SenderGender");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SenderGender", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.LetterContent.SenderReference"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string SenderReference
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SenderReference");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SenderReference", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
