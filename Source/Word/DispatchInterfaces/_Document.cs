using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface _Document 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Name"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Application"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Creator"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Parent"/> </remarks>
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
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.BuiltInDocumentProperties"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public object BuiltInDocumentProperties
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "BuiltInDocumentProperties");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CustomDocumentProperties"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public object CustomDocumentProperties
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "CustomDocumentProperties");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Path"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string Path
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Path");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Bookmarks"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Bookmarks Bookmarks
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Bookmarks>(this, "Bookmarks", NetOffice.WordApi.Bookmarks.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Tables"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Tables Tables
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Tables>(this, "Tables", NetOffice.WordApi.Tables.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Footnotes"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Footnotes Footnotes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Footnotes>(this, "Footnotes", NetOffice.WordApi.Footnotes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Endnotes"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Endnotes Endnotes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Endnotes>(this, "Endnotes", NetOffice.WordApi.Endnotes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Comments"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Comments Comments
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Comments>(this, "Comments", NetOffice.WordApi.Comments.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Type"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdDocumentType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdDocumentType>(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.AutoHyphenation"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoHyphenation
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoHyphenation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoHyphenation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.HyphenateCaps"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool HyphenateCaps
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HyphenateCaps");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HyphenateCaps", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.HyphenationZone"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 HyphenationZone
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "HyphenationZone");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HyphenationZone", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ConsecutiveHyphensLimit"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 ConsecutiveHyphensLimit
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ConsecutiveHyphensLimit");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ConsecutiveHyphensLimit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Sections"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Sections Sections
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Sections>(this, "Sections", NetOffice.WordApi.Sections.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Paragraphs"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Paragraphs Paragraphs
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Paragraphs>(this, "Paragraphs", NetOffice.WordApi.Paragraphs.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Words"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Words Words
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Words>(this, "Words", NetOffice.WordApi.Words.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Sentences"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Sentences Sentences
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Sentences>(this, "Sentences", NetOffice.WordApi.Sentences.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Characters"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Characters Characters
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Characters>(this, "Characters", NetOffice.WordApi.Characters.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Fields"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Fields Fields
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Fields>(this, "Fields", NetOffice.WordApi.Fields.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.FormFields"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.FormFields FormFields
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.FormFields>(this, "FormFields", NetOffice.WordApi.FormFields.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Styles"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Styles Styles
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Styles>(this, "Styles", NetOffice.WordApi.Styles.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Frames"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Frames Frames
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Frames>(this, "Frames", NetOffice.WordApi.Frames.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.TablesOfFigures"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TablesOfFigures TablesOfFigures
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.TablesOfFigures>(this, "TablesOfFigures", NetOffice.WordApi.TablesOfFigures.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Variables"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Variables Variables
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Variables>(this, "Variables", NetOffice.WordApi.Variables.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.MailMerge"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMerge MailMerge
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.MailMerge>(this, "MailMerge", NetOffice.WordApi.MailMerge.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Envelope"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Envelope Envelope
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Envelope>(this, "Envelope", NetOffice.WordApi.Envelope.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.FullName"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string FullName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FullName");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Revisions"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Revisions Revisions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Revisions>(this, "Revisions", NetOffice.WordApi.Revisions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.TablesOfContents"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TablesOfContents TablesOfContents
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.TablesOfContents>(this, "TablesOfContents", NetOffice.WordApi.TablesOfContents.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.TablesOfAuthorities"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TablesOfAuthorities TablesOfAuthorities
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.TablesOfAuthorities>(this, "TablesOfAuthorities", NetOffice.WordApi.TablesOfAuthorities.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PageSetup"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.PageSetup PageSetup
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.PageSetup>(this, "PageSetup", NetOffice.WordApi.PageSetup.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "PageSetup", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Windows"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Windows Windows
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Windows>(this, "Windows", NetOffice.WordApi.Windows.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool HasRoutingSlip
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasRoutingSlip");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasRoutingSlip", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.RoutingSlip RoutingSlip
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.RoutingSlip>(this, "RoutingSlip", NetOffice.WordApi.RoutingSlip.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Routed
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Routed");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.TablesOfAuthoritiesCategories"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TablesOfAuthoritiesCategories TablesOfAuthoritiesCategories
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.TablesOfAuthoritiesCategories>(this, "TablesOfAuthoritiesCategories", NetOffice.WordApi.TablesOfAuthoritiesCategories.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Indexes"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Indexes Indexes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Indexes>(this, "Indexes", NetOffice.WordApi.Indexes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Saved"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Saved
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Saved");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Saved", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Content"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Content
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "Content", NetOffice.WordApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ActiveWindow"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Window ActiveWindow
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Window>(this, "ActiveWindow", NetOffice.WordApi.Window.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Kind"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdDocumentKind Kind
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdDocumentKind>(this, "Kind");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Kind", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ReadOnly"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ReadOnly
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ReadOnly");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Subdocuments"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Subdocuments Subdocuments
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Subdocuments>(this, "Subdocuments", NetOffice.WordApi.Subdocuments.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.IsMasterDocument"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool IsMasterDocument
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsMasterDocument");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.DefaultTabStop"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single DefaultTabStop
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "DefaultTabStop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultTabStop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.EmbedTrueTypeFonts"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool EmbedTrueTypeFonts
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EmbedTrueTypeFonts");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EmbedTrueTypeFonts", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SaveFormsData"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool SaveFormsData
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SaveFormsData");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SaveFormsData", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ReadOnlyRecommended"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ReadOnlyRecommended
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ReadOnlyRecommended");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReadOnlyRecommended", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SaveSubsetFonts"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool SaveSubsetFonts
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SaveSubsetFonts");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SaveSubsetFonts", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Compatibility"/> </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdCompatibility type</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool get_Compatibility(NetOffice.WordApi.Enums.WdCompatibility type)
		{
			return Factory.ExecuteBoolPropertyGet(this, "Compatibility", type);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdCompatibility type</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_Compatibility(NetOffice.WordApi.Enums.WdCompatibility type, bool value)
		{
			Factory.ExecutePropertySet(this, "Compatibility", type, value);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Compatibility
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Compatibility"/> </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdCompatibility type</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), Redirect("get_Compatibility")]
		public bool Compatibility(NetOffice.WordApi.Enums.WdCompatibility type)
		{
			return get_Compatibility(type);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.StoryRanges"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.StoryRanges StoryRanges
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.StoryRanges>(this, "StoryRanges", NetOffice.WordApi.StoryRanges.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CommandBars"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBars CommandBars
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBars>(this, "CommandBars", NetOffice.OfficeApi.CommandBars.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.IsSubdocument"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool IsSubdocument
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsSubdocument");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SaveFormat"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 SaveFormat
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SaveFormat");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ProtectionType"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdProtectionType ProtectionType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdProtectionType>(this, "ProtectionType");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Hyperlinks"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Hyperlinks Hyperlinks
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Hyperlinks>(this, "Hyperlinks", NetOffice.WordApi.Hyperlinks.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Shapes"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shapes Shapes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Shapes>(this, "Shapes", NetOffice.WordApi.Shapes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ListTemplates"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ListTemplates ListTemplates
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ListTemplates>(this, "ListTemplates", NetOffice.WordApi.ListTemplates.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Lists"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Lists Lists
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Lists>(this, "Lists", NetOffice.WordApi.Lists.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.UpdateStylesOnOpen"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool UpdateStylesOnOpen
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UpdateStylesOnOpen");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UpdateStylesOnOpen", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.AttachedTemplate"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object AttachedTemplate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "AttachedTemplate");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "AttachedTemplate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.InlineShapes"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShapes InlineShapes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.InlineShapes>(this, "InlineShapes", NetOffice.WordApi.InlineShapes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Background"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape Background
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Shape>(this, "Background", NetOffice.WordApi.Shape.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Background", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.GrammarChecked"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool GrammarChecked
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "GrammarChecked");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GrammarChecked", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SpellingChecked"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool SpellingChecked
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SpellingChecked");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SpellingChecked", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ShowGrammaticalErrors"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShowGrammaticalErrors
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowGrammaticalErrors");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowGrammaticalErrors", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ShowSpellingErrors"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShowSpellingErrors
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowSpellingErrors");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowSpellingErrors", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Versions Versions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Versions>(this, "Versions", NetOffice.WordApi.Versions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShowSummary
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowSummary");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowSummary", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdSummaryMode SummaryViewMode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdSummaryMode>(this, "SummaryViewMode");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SummaryViewMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 SummaryLength
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SummaryLength");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SummaryLength", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool PrintFractionalWidths
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintFractionalWidths");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintFractionalWidths", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PrintPostScriptOverText"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool PrintPostScriptOverText
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintPostScriptOverText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintPostScriptOverText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Container"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public object Container
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Container");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PrintFormsData"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool PrintFormsData
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintFormsData");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintFormsData", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ListParagraphs"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ListParagraphs ListParagraphs
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ListParagraphs>(this, "ListParagraphs", NetOffice.WordApi.ListParagraphs.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Password"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string Password
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Password");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Password", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.WritePassword"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string WritePassword
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "WritePassword");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WritePassword", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.HasPassword"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool HasPassword
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasPassword");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.WriteReserved"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool WriteReserved
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "WriteReserved");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ActiveWritingStyle"/> </remarks>
		/// <param name="languageID">object languageID</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_ActiveWritingStyle(object languageID)
		{
			return Factory.ExecuteStringPropertyGet(this, "ActiveWritingStyle", languageID);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="languageID">object languageID</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_ActiveWritingStyle(object languageID, string value)
		{
			Factory.ExecutePropertySet(this, "ActiveWritingStyle", languageID, value);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_ActiveWritingStyle
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ActiveWritingStyle"/> </remarks>
		/// <param name="languageID">object languageID</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), Redirect("get_ActiveWritingStyle")]
		public string ActiveWritingStyle(object languageID)
		{
			return get_ActiveWritingStyle(languageID);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.UserControl"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool UserControl
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UserControl");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UserControl", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool HasMailer
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasMailer");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasMailer", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Mailer Mailer
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Mailer>(this, "Mailer", NetOffice.WordApi.Mailer.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ReadabilityStatistics"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ReadabilityStatistics ReadabilityStatistics
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ReadabilityStatistics>(this, "ReadabilityStatistics", NetOffice.WordApi.ReadabilityStatistics.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.GrammaticalErrors"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ProofreadingErrors GrammaticalErrors
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ProofreadingErrors>(this, "GrammaticalErrors", NetOffice.WordApi.ProofreadingErrors.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SpellingErrors"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ProofreadingErrors SpellingErrors
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ProofreadingErrors>(this, "SpellingErrors", NetOffice.WordApi.ProofreadingErrors.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.VBProject"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.VBIDEApi.VBProject VBProject
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBProject>(this, "VBProject", NetOffice.VBIDEApi.VBProject.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.FormsDesign"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool FormsDesign
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FormsDesign");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string _CodeName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "_CodeName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "_CodeName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CodeName"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string CodeName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CodeName");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SnapToGrid"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool SnapToGrid
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SnapToGrid");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SnapToGrid", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SnapToShapes"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool SnapToShapes
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SnapToShapes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SnapToShapes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.GridDistanceHorizontal"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single GridDistanceHorizontal
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "GridDistanceHorizontal");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridDistanceHorizontal", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.GridDistanceVertical"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single GridDistanceVertical
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "GridDistanceVertical");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridDistanceVertical", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.GridOriginHorizontal"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single GridOriginHorizontal
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "GridOriginHorizontal");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridOriginHorizontal", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.GridOriginVertical"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single GridOriginVertical
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "GridOriginVertical");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridOriginVertical", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.GridSpaceBetweenHorizontalLines"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 GridSpaceBetweenHorizontalLines
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "GridSpaceBetweenHorizontalLines");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridSpaceBetweenHorizontalLines", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.GridSpaceBetweenVerticalLines"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 GridSpaceBetweenVerticalLines
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "GridSpaceBetweenVerticalLines");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridSpaceBetweenVerticalLines", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.GridOriginFromMargin"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool GridOriginFromMargin
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "GridOriginFromMargin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridOriginFromMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.KerningByAlgorithm"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool KerningByAlgorithm
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "KerningByAlgorithm");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "KerningByAlgorithm", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.JustificationMode"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdJustificationMode JustificationMode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdJustificationMode>(this, "JustificationMode");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "JustificationMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.FarEastLineBreakLevel"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdFarEastLineBreakLevel FarEastLineBreakLevel
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdFarEastLineBreakLevel>(this, "FarEastLineBreakLevel");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "FarEastLineBreakLevel", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.NoLineBreakBefore"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string NoLineBreakBefore
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NoLineBreakBefore");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NoLineBreakBefore", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.NoLineBreakAfter"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string NoLineBreakAfter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NoLineBreakAfter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NoLineBreakAfter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.TrackRevisions"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool TrackRevisions
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TrackRevisions");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TrackRevisions", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PrintRevisions"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool PrintRevisions
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintRevisions");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintRevisions", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShowRevisions
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowRevisions");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowRevisions", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ActiveTheme"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string ActiveTheme
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ActiveTheme");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ActiveThemeDisplayName"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string ActiveThemeDisplayName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ActiveThemeDisplayName");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Email"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Email Email
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Email>(this, "Email", NetOffice.WordApi.Email.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Scripts"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Scripts Scripts
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Scripts>(this, "Scripts", NetOffice.OfficeApi.Scripts.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.LanguageDetected"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool LanguageDetected
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "LanguageDetected");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LanguageDetected", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.FarEastLineBreakLanguage"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdFarEastLineBreakLanguageID FarEastLineBreakLanguage
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdFarEastLineBreakLanguageID>(this, "FarEastLineBreakLanguage");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "FarEastLineBreakLanguage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Frameset"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Frameset Frameset
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Frameset>(this, "Frameset", NetOffice.WordApi.Frameset.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ClickAndTypeParagraphStyle"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object ClickAndTypeParagraphStyle
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ClickAndTypeParagraphStyle");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ClickAndTypeParagraphStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.HTMLProject HTMLProject
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.HTMLProject>(this, "HTMLProject", NetOffice.OfficeApi.HTMLProject.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.WebOptions"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.WebOptions WebOptions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.WebOptions>(this, "WebOptions", NetOffice.WordApi.WebOptions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.OpenEncoding"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoEncoding OpenEncoding
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoEncoding>(this, "OpenEncoding");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SaveEncoding"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoEncoding SaveEncoding
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoEncoding>(this, "SaveEncoding");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SaveEncoding", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.OptimizeForWord97"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool OptimizeForWord97
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "OptimizeForWord97");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OptimizeForWord97", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.VBASigned"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool VBASigned
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "VBASigned");
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.MailEnvelope"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.MsoEnvelope MailEnvelope
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.MsoEnvelope>(this, "MailEnvelope", NetOffice.OfficeApi.MsoEnvelope.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.DisableFeatures"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool DisableFeatures
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisableFeatures");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisableFeatures", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.DoNotEmbedSystemFonts"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool DoNotEmbedSystemFonts
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DoNotEmbedSystemFonts");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DoNotEmbedSystemFonts", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Signatures"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.SignatureSet Signatures
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SignatureSet>(this, "Signatures", NetOffice.OfficeApi.SignatureSet.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.DefaultTargetFrame"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public string DefaultTargetFrame
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DefaultTargetFrame");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultTargetFrame", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.HTMLDivisions"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.HTMLDivisions HTMLDivisions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.HTMLDivisions>(this, "HTMLDivisions", NetOffice.WordApi.HTMLDivisions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.DisableFeaturesIntroducedAfter"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdDisableFeaturesIntroducedAfter DisableFeaturesIntroducedAfter
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdDisableFeaturesIntroducedAfter>(this, "DisableFeaturesIntroducedAfter");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DisableFeaturesIntroducedAfter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.RemovePersonalInformation"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
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
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.SmartTags SmartTags
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.SmartTags>(this, "SmartTags", NetOffice.WordApi.SmartTags.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool EmbedSmartTags
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EmbedSmartTags");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EmbedSmartTags", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool SmartTagsAsXMLProps
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SmartTagsAsXMLProps");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SmartTagsAsXMLProps", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.TextEncoding"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoEncoding TextEncoding
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoEncoding>(this, "TextEncoding");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TextEncoding", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.TextLineEnding"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdLineEndingType TextLineEnding
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdLineEndingType>(this, "TextLineEnding");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TextLineEnding", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.StyleSheets"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.StyleSheets StyleSheets
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.StyleSheets>(this, "StyleSheets", NetOffice.WordApi.StyleSheets.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.DefaultTableStyle"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public object DefaultTableStyle
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DefaultTableStyle");
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PasswordEncryptionProvider"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public string PasswordEncryptionProvider
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PasswordEncryptionProvider");
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PasswordEncryptionAlgorithm"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public string PasswordEncryptionAlgorithm
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PasswordEncryptionAlgorithm");
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PasswordEncryptionKeyLength"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public Int32 PasswordEncryptionKeyLength
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PasswordEncryptionKeyLength");
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PasswordEncryptionFileProperties"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool PasswordEncryptionFileProperties
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PasswordEncryptionFileProperties");
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.EmbedLinguisticData"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool EmbedLinguisticData
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EmbedLinguisticData");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EmbedLinguisticData", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.FormattingShowFont"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool FormattingShowFont
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FormattingShowFont");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FormattingShowFont", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.FormattingShowClear"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool FormattingShowClear
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FormattingShowClear");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FormattingShowClear", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.FormattingShowParagraph"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool FormattingShowParagraph
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FormattingShowParagraph");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FormattingShowParagraph", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.FormattingShowNumbering"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool FormattingShowNumbering
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FormattingShowNumbering");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FormattingShowNumbering", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.FormattingShowFilter"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdShowFilter FormattingShowFilter
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdShowFilter>(this, "FormattingShowFilter");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "FormattingShowFilter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Permission"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.OfficeApi.Permission Permission
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Permission>(this, "Permission", NetOffice.OfficeApi.Permission.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNodes XMLNodes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLNodes>(this, "XMLNodes", NetOffice.WordApi.XMLNodes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.XMLSchemaReferences"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLSchemaReferences XMLSchemaReferences
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLSchemaReferences>(this, "XMLSchemaReferences", NetOffice.WordApi.XMLSchemaReferences.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SmartDocument"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.OfficeApi.SmartDocument SmartDocument
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartDocument>(this, "SmartDocument", NetOffice.OfficeApi.SmartDocument.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.OfficeApi.SharedWorkspace SharedWorkspace
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SharedWorkspace>(this, "SharedWorkspace", NetOffice.OfficeApi.SharedWorkspace.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Sync(property)"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.OfficeApi.Sync Sync
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Sync>(this, "Sync", NetOffice.OfficeApi.Sync.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.EnforceStyle"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public bool EnforceStyle
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnforceStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnforceStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.AutoFormatOverride"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public bool AutoFormatOverride
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatOverride");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatOverride", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public bool XMLSaveDataOnly
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "XMLSaveDataOnly");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "XMLSaveDataOnly", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public bool XMLHideNamespaces
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "XMLHideNamespaces");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "XMLHideNamespaces", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.XMLShowAdvancedErrors"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public bool XMLShowAdvancedErrors
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "XMLShowAdvancedErrors");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "XMLShowAdvancedErrors", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.XMLUseXSLTWhenSaving"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public bool XMLUseXSLTWhenSaving
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "XMLUseXSLTWhenSaving");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "XMLUseXSLTWhenSaving", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.XMLSaveThroughXSLT"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public string XMLSaveThroughXSLT
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "XMLSaveThroughXSLT");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "XMLSaveThroughXSLT", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.DocumentLibraryVersions"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.OfficeApi.DocumentLibraryVersions DocumentLibraryVersions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.DocumentLibraryVersions>(this, "DocumentLibraryVersions", NetOffice.OfficeApi.DocumentLibraryVersions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ReadingModeLayoutFrozen"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public bool ReadingModeLayoutFrozen
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ReadingModeLayoutFrozen");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReadingModeLayoutFrozen", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.RemoveDateAndTime"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public bool RemoveDateAndTime
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RemoveDateAndTime");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RemoveDateAndTime", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLChildNodeSuggestions ChildNodeSuggestions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLChildNodeSuggestions>(this, "ChildNodeSuggestions", NetOffice.WordApi.XMLChildNodeSuggestions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNodes XMLSchemaViolations
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLNodes>(this, "XMLSchemaViolations", NetOffice.WordApi.XMLNodes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ReadingLayoutSizeX"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public Int32 ReadingLayoutSizeX
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ReadingLayoutSizeX");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReadingLayoutSizeX", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ReadingLayoutSizeY"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public Int32 ReadingLayoutSizeY
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ReadingLayoutSizeY");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReadingLayoutSizeY", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.StyleSortMethod"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdStyleSort StyleSortMethod
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdStyleSort>(this, "StyleSortMethod");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "StyleSortMethod", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ContentTypeProperties"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.OfficeApi.MetaProperties ContentTypeProperties
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.MetaProperties>(this, "ContentTypeProperties", NetOffice.OfficeApi.MetaProperties.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.TrackMoves"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool TrackMoves
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TrackMoves");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TrackMoves", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.TrackFormatting"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool TrackFormatting
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TrackFormatting");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TrackFormatting", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Dummy1
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Dummy1");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.OMaths"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMaths OMaths
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMaths>(this, "OMaths", NetOffice.WordApi.OMaths.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Dummy3
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Dummy3");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ServerPolicy"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.OfficeApi.ServerPolicy ServerPolicy
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ServerPolicy>(this, "ServerPolicy", NetOffice.OfficeApi.ServerPolicy.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ContentControls"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.ContentControls ContentControls
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ContentControls>(this, "ContentControls", NetOffice.WordApi.ContentControls.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.DocumentInspectors"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.OfficeApi.DocumentInspectors DocumentInspectors
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.DocumentInspectors>(this, "DocumentInspectors", NetOffice.OfficeApi.DocumentInspectors.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Bibliography"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Bibliography Bibliography
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Bibliography>(this, "Bibliography", NetOffice.WordApi.Bibliography.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.LockTheme"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool LockTheme
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "LockTheme");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LockTheme", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.LockQuickStyleSet"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool LockQuickStyleSet
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "LockQuickStyleSet");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LockQuickStyleSet", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.OriginalDocumentTitle"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public string OriginalDocumentTitle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OriginalDocumentTitle");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.RevisedDocumentTitle"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public string RevisedDocumentTitle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RevisedDocumentTitle");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CustomXMLParts"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLParts CustomXMLParts
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLParts>(this, "CustomXMLParts", NetOffice.OfficeApi.CustomXMLParts.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.FormattingShowNextLevel"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool FormattingShowNextLevel
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FormattingShowNextLevel");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FormattingShowNextLevel", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.FormattingShowUserStyleName"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool FormattingShowUserStyleName
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FormattingShowUserStyleName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FormattingShowUserStyleName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Research"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Research Research
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Research>(this, "Research", NetOffice.WordApi.Research.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Final"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool Final
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Final");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Final", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.OMathBreakBin"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdOMathBreakBin OMathBreakBin
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdOMathBreakBin>(this, "OMathBreakBin");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "OMathBreakBin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.OMathBreakSub"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdOMathBreakSub OMathBreakSub
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdOMathBreakSub>(this, "OMathBreakSub");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "OMathBreakSub", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.OMathJc"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdOMathJc OMathJc
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdOMathJc>(this, "OMathJc");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "OMathJc", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.OMathLeftMargin"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public Single OMathLeftMargin
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "OMathLeftMargin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OMathLeftMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.OMathRightMargin"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public Single OMathRightMargin
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "OMathRightMargin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OMathRightMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.OMathWrap"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public Single OMathWrap
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "OMathWrap");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OMathWrap", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.OMathIntSubSupLim"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool OMathIntSubSupLim
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "OMathIntSubSupLim");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OMathIntSubSupLim", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.OMathNarySupSubLim"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool OMathNarySupSubLim
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "OMathNarySupSubLim");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OMathNarySupSubLim", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.OMathSmallFrac"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool OMathSmallFrac
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "OMathSmallFrac");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OMathSmallFrac", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.WordOpenXML"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public string WordOpenXML
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "WordOpenXML");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.DocumentTheme"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.OfficeApi.OfficeTheme DocumentTheme
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.OfficeTheme>(this, "DocumentTheme", NetOffice.OfficeApi.OfficeTheme.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.HasVBProject"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool HasVBProject
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasVBProject");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.OMathFontName"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public string OMathFontName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OMathFontName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OMathFontName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.EncryptionProvider"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public string EncryptionProvider
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EncryptionProvider");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EncryptionProvider", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.UseMathDefaults"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool UseMathDefaults
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseMathDefaults");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseMathDefaults", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CurrentRsid"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public Int32 CurrentRsid
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CurrentRsid");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 DocID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DocID");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CompatibilityMode"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public Int32 CompatibilityMode
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CompatibilityMode");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CoAuthoring"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.CoAuthoring CoAuthoring
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.CoAuthoring>(this, "CoAuthoring", NetOffice.WordApi.CoAuthoring.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.document.broadcast"/> </remarks>
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Broadcast Broadcast
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Broadcast>(this, "Broadcast", NetOffice.WordApi.Broadcast.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.document.chartdatapointtrack"/> </remarks>
		[SupportByVersion("Word", 15, 16)]
		public bool ChartDataPointTrack
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ChartDataPointTrack");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ChartDataPointTrack", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.document.isinautosave"/> </remarks>
		[SupportByVersion("Word", 15, 16)]
		public bool IsInAutosave
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsInAutosave");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Close(method)"/> </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		/// <param name="originalFormat">optional object originalFormat</param>
		/// <param name="routeDocument">optional object routeDocument</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Close(object saveChanges, object originalFormat, object routeDocument)
		{
			 Factory.ExecuteMethod(this, "Close", saveChanges, originalFormat, routeDocument);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Close(method)"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Close()
		{
			 Factory.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Close(method)"/> </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Close(object saveChanges)
		{
			 Factory.ExecuteMethod(this, "Close", saveChanges);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Close(method)"/> </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		/// <param name="originalFormat">optional object originalFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Close(object saveChanges, object originalFormat)
		{
			 Factory.ExecuteMethod(this, "Close", saveChanges, originalFormat);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object allowSubstitutions</param>
		/// <param name="lineEnding">optional object lineEnding</param>
		/// <param name="addBiDiMarks">optional object addBiDiMarks</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding, object addBiDiMarks)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions, lineEnding, addBiDiMarks });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SaveAs()
		{
			 Factory.ExecuteMethod(this, "SaveAs");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SaveAs(object fileName)
		{
			 Factory.ExecuteMethod(this, "SaveAs", fileName);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat)
		{
			 Factory.ExecuteMethod(this, "SaveAs", fileName, fileFormat);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments)
		{
			 Factory.ExecuteMethod(this, "SaveAs", fileName, fileFormat, lockComments);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments, object password)
		{
			 Factory.ExecuteMethod(this, "SaveAs", fileName, fileFormat, lockComments, password);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object allowSubstitutions</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object allowSubstitutions</param>
		/// <param name="lineEnding">optional object lineEnding</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions, lineEnding });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Repaginate"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Repaginate()
		{
			 Factory.ExecuteMethod(this, "Repaginate");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.FitToPages"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void FitToPages()
		{
			 Factory.ExecuteMethod(this, "FitToPages");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ManualHyphenation"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ManualHyphenation()
		{
			 Factory.ExecuteMethod(this, "ManualHyphenation");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Select"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Select()
		{
			 Factory.ExecuteMethod(this, "Select");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.DataForm"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void DataForm()
		{
			 Factory.ExecuteMethod(this, "DataForm");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Route()
		{
			 Factory.ExecuteMethod(this, "Route");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Save"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Save()
		{
			 Factory.ExecuteMethod(this, "Save");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld()
		{
			 Factory.ExecuteMethod(this, "PrintOutOld");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", background);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", background, append);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", background, append, range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", background, append, range, outputFileName);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SendMail"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SendMail()
		{
			 Factory.ExecuteMethod(this, "SendMail");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Range"/> </remarks>
		/// <param name="start">optional object start</param>
		/// <param name="end">optional object end</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Range(object start, object end)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "Range", NetOffice.WordApi.Range.LateBindingApiWrapperType, start, end);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Range"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Range()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "Range", NetOffice.WordApi.Range.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Range"/> </remarks>
		/// <param name="start">optional object start</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Range(object start)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "Range", NetOffice.WordApi.Range.LateBindingApiWrapperType, start);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.RunAutoMacro"/> </remarks>
		/// <param name="which">NetOffice.WordApi.Enums.WdAutoMacros which</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void RunAutoMacro(NetOffice.WordApi.Enums.WdAutoMacros which)
		{
			 Factory.ExecuteMethod(this, "RunAutoMacro", which);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Activate"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Activate()
		{
			 Factory.ExecuteMethod(this, "Activate");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PrintPreview"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintPreview()
		{
			 Factory.ExecuteMethod(this, "PrintPreview");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.GoTo"/> </remarks>
		/// <param name="what">optional object what</param>
		/// <param name="which">optional object which</param>
		/// <param name="count">optional object count</param>
		/// <param name="name">optional object name</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range GoTo(object what, object which, object count, object name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoTo", NetOffice.WordApi.Range.LateBindingApiWrapperType, what, which, count, name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.GoTo"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range GoTo()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoTo", NetOffice.WordApi.Range.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.GoTo"/> </remarks>
		/// <param name="what">optional object what</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range GoTo(object what)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoTo", NetOffice.WordApi.Range.LateBindingApiWrapperType, what);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.GoTo"/> </remarks>
		/// <param name="what">optional object what</param>
		/// <param name="which">optional object which</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range GoTo(object what, object which)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoTo", NetOffice.WordApi.Range.LateBindingApiWrapperType, what, which);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.GoTo"/> </remarks>
		/// <param name="what">optional object what</param>
		/// <param name="which">optional object which</param>
		/// <param name="count">optional object count</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range GoTo(object what, object which, object count)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoTo", NetOffice.WordApi.Range.LateBindingApiWrapperType, what, which, count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Undo"/> </remarks>
		/// <param name="times">optional object times</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Undo(object times)
		{
			return Factory.ExecuteBoolMethodGet(this, "Undo", times);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Undo"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Undo()
		{
			return Factory.ExecuteBoolMethodGet(this, "Undo");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Redo"/> </remarks>
		/// <param name="times">optional object times</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Redo(object times)
		{
			return Factory.ExecuteBoolMethodGet(this, "Redo", times);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Redo"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Redo()
		{
			return Factory.ExecuteBoolMethodGet(this, "Redo");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ComputeStatistics"/> </remarks>
		/// <param name="statistic">NetOffice.WordApi.Enums.WdStatistic statistic</param>
		/// <param name="includeFootnotesAndEndnotes">optional object includeFootnotesAndEndnotes</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 ComputeStatistics(NetOffice.WordApi.Enums.WdStatistic statistic, object includeFootnotesAndEndnotes)
		{
			return Factory.ExecuteInt32MethodGet(this, "ComputeStatistics", statistic, includeFootnotesAndEndnotes);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ComputeStatistics"/> </remarks>
		/// <param name="statistic">NetOffice.WordApi.Enums.WdStatistic statistic</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 ComputeStatistics(NetOffice.WordApi.Enums.WdStatistic statistic)
		{
			return Factory.ExecuteInt32MethodGet(this, "ComputeStatistics", statistic);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.MakeCompatibilityDefault"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void MakeCompatibilityDefault()
		{
			 Factory.ExecuteMethod(this, "MakeCompatibilityDefault");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.document.protect"/> </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType type</param>
		/// <param name="noReset">optional object noReset</param>
		/// <param name="password">optional object password</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Protect(NetOffice.WordApi.Enums.WdProtectionType type, object noReset, object password)
		{
			 Factory.ExecuteMethod(this, "Protect", type, noReset, password);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.document.protect"/> </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType type</param>
		/// <param name="noReset">optional object noReset</param>
		/// <param name="password">optional object password</param>
		/// <param name="useIRM">optional object useIRM</param>
		/// <param name="enforceStyleLock">optional object enforceStyleLock</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void Protect(NetOffice.WordApi.Enums.WdProtectionType type, object noReset, object password, object useIRM, object enforceStyleLock)
		{
			 Factory.ExecuteMethod(this, "Protect", new object[]{ type, noReset, password, useIRM, enforceStyleLock });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.document.protect"/> </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType type</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Protect(NetOffice.WordApi.Enums.WdProtectionType type)
		{
			 Factory.ExecuteMethod(this, "Protect", type);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.document.protect"/> </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType type</param>
		/// <param name="noReset">optional object noReset</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Protect(NetOffice.WordApi.Enums.WdProtectionType type, object noReset)
		{
			 Factory.ExecuteMethod(this, "Protect", type, noReset);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.document.protect"/> </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType type</param>
		/// <param name="noReset">optional object noReset</param>
		/// <param name="password">optional object password</param>
		/// <param name="useIRM">optional object useIRM</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void Protect(NetOffice.WordApi.Enums.WdProtectionType type, object noReset, object password, object useIRM)
		{
			 Factory.ExecuteMethod(this, "Protect", type, noReset, password, useIRM);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Unprotect"/> </remarks>
		/// <param name="password">optional object password</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Unprotect(object password)
		{
			 Factory.ExecuteMethod(this, "Unprotect", password);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Unprotect"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Unprotect()
		{
			 Factory.ExecuteMethod(this, "Unprotect");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdEditionType type</param>
		/// <param name="option">NetOffice.WordApi.Enums.WdEditionOption option</param>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void EditionOptions(NetOffice.WordApi.Enums.WdEditionType type, NetOffice.WordApi.Enums.WdEditionOption option, string name, object format)
		{
			 Factory.ExecuteMethod(this, "EditionOptions", type, option, name, format);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdEditionType type</param>
		/// <param name="option">NetOffice.WordApi.Enums.WdEditionOption option</param>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void EditionOptions(NetOffice.WordApi.Enums.WdEditionType type, NetOffice.WordApi.Enums.WdEditionOption option, string name)
		{
			 Factory.ExecuteMethod(this, "EditionOptions", type, option, name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.RunLetterWizard"/> </remarks>
		/// <param name="letterContent">optional object letterContent</param>
		/// <param name="wizardMode">optional object wizardMode</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void RunLetterWizard(object letterContent, object wizardMode)
		{
			 Factory.ExecuteMethod(this, "RunLetterWizard", letterContent, wizardMode);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.RunLetterWizard"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void RunLetterWizard()
		{
			 Factory.ExecuteMethod(this, "RunLetterWizard");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.RunLetterWizard"/> </remarks>
		/// <param name="letterContent">optional object letterContent</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void RunLetterWizard(object letterContent)
		{
			 Factory.ExecuteMethod(this, "RunLetterWizard", letterContent);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.GetLetterContent"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.LetterContent GetLetterContent()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.LetterContent>(this, "GetLetterContent", NetOffice.WordApi.LetterContent.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SetLetterContent"/> </remarks>
		/// <param name="letterContent">object letterContent</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SetLetterContent(object letterContent)
		{
			 Factory.ExecuteMethod(this, "SetLetterContent", letterContent);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CopyStylesFromTemplate"/> </remarks>
		/// <param name="template">string template</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CopyStylesFromTemplate(string template)
		{
			 Factory.ExecuteMethod(this, "CopyStylesFromTemplate", template);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.UpdateStyles"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void UpdateStyles()
		{
			 Factory.ExecuteMethod(this, "UpdateStyles");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CheckGrammar"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckGrammar()
		{
			 Factory.ExecuteMethod(this, "CheckGrammar");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CheckSpelling"/> </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		/// <param name="customDictionary6">optional object customDictionary6</param>
		/// <param name="customDictionary7">optional object customDictionary7</param>
		/// <param name="customDictionary8">optional object customDictionary8</param>
		/// <param name="customDictionary9">optional object customDictionary9</param>
		/// <param name="customDictionary10">optional object customDictionary10</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9, object customDictionary10)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9, customDictionary10 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CheckSpelling"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling()
		{
			 Factory.ExecuteMethod(this, "CheckSpelling");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CheckSpelling"/> </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", customDictionary);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CheckSpelling"/> </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", customDictionary, ignoreUppercase);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CheckSpelling"/> </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", customDictionary, ignoreUppercase, alwaysSuggest);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CheckSpelling"/> </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CheckSpelling"/> </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CheckSpelling"/> </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CheckSpelling"/> </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CheckSpelling"/> </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		/// <param name="customDictionary6">optional object customDictionary6</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CheckSpelling"/> </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		/// <param name="customDictionary6">optional object customDictionary6</param>
		/// <param name="customDictionary7">optional object customDictionary7</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CheckSpelling"/> </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		/// <param name="customDictionary6">optional object customDictionary6</param>
		/// <param name="customDictionary7">optional object customDictionary7</param>
		/// <param name="customDictionary8">optional object customDictionary8</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CheckSpelling"/> </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		/// <param name="customDictionary6">optional object customDictionary6</param>
		/// <param name="customDictionary7">optional object customDictionary7</param>
		/// <param name="customDictionary8">optional object customDictionary8</param>
		/// <param name="customDictionary9">optional object customDictionary9</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.FollowHyperlink"/> </remarks>
		/// <param name="address">optional object address</param>
		/// <param name="subAddress">optional object subAddress</param>
		/// <param name="newWindow">optional object newWindow</param>
		/// <param name="addHistory">optional object addHistory</param>
		/// <param name="extraInfo">optional object extraInfo</param>
		/// <param name="method">optional object method</param>
		/// <param name="headerInfo">optional object headerInfo</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(object address, object subAddress, object newWindow, object addHistory, object extraInfo, object method, object headerInfo)
		{
			 Factory.ExecuteMethod(this, "FollowHyperlink", new object[]{ address, subAddress, newWindow, addHistory, extraInfo, method, headerInfo });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.FollowHyperlink"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void FollowHyperlink()
		{
			 Factory.ExecuteMethod(this, "FollowHyperlink");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.FollowHyperlink"/> </remarks>
		/// <param name="address">optional object address</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(object address)
		{
			 Factory.ExecuteMethod(this, "FollowHyperlink", address);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.FollowHyperlink"/> </remarks>
		/// <param name="address">optional object address</param>
		/// <param name="subAddress">optional object subAddress</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(object address, object subAddress)
		{
			 Factory.ExecuteMethod(this, "FollowHyperlink", address, subAddress);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.FollowHyperlink"/> </remarks>
		/// <param name="address">optional object address</param>
		/// <param name="subAddress">optional object subAddress</param>
		/// <param name="newWindow">optional object newWindow</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(object address, object subAddress, object newWindow)
		{
			 Factory.ExecuteMethod(this, "FollowHyperlink", address, subAddress, newWindow);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.FollowHyperlink"/> </remarks>
		/// <param name="address">optional object address</param>
		/// <param name="subAddress">optional object subAddress</param>
		/// <param name="newWindow">optional object newWindow</param>
		/// <param name="addHistory">optional object addHistory</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(object address, object subAddress, object newWindow, object addHistory)
		{
			 Factory.ExecuteMethod(this, "FollowHyperlink", address, subAddress, newWindow, addHistory);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.FollowHyperlink"/> </remarks>
		/// <param name="address">optional object address</param>
		/// <param name="subAddress">optional object subAddress</param>
		/// <param name="newWindow">optional object newWindow</param>
		/// <param name="addHistory">optional object addHistory</param>
		/// <param name="extraInfo">optional object extraInfo</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(object address, object subAddress, object newWindow, object addHistory, object extraInfo)
		{
			 Factory.ExecuteMethod(this, "FollowHyperlink", new object[]{ address, subAddress, newWindow, addHistory, extraInfo });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.FollowHyperlink"/> </remarks>
		/// <param name="address">optional object address</param>
		/// <param name="subAddress">optional object subAddress</param>
		/// <param name="newWindow">optional object newWindow</param>
		/// <param name="addHistory">optional object addHistory</param>
		/// <param name="extraInfo">optional object extraInfo</param>
		/// <param name="method">optional object method</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(object address, object subAddress, object newWindow, object addHistory, object extraInfo, object method)
		{
			 Factory.ExecuteMethod(this, "FollowHyperlink", new object[]{ address, subAddress, newWindow, addHistory, extraInfo, method });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.AddToFavorites"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void AddToFavorites()
		{
			 Factory.ExecuteMethod(this, "AddToFavorites");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Reload"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Reload()
		{
			 Factory.ExecuteMethod(this, "Reload");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="length">optional object length</param>
		/// <param name="mode">optional object mode</param>
		/// <param name="updateProperties">optional object updateProperties</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range AutoSummarize(object length, object mode, object updateProperties)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "AutoSummarize", NetOffice.WordApi.Range.LateBindingApiWrapperType, length, mode, updateProperties);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range AutoSummarize()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "AutoSummarize", NetOffice.WordApi.Range.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="length">optional object length</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range AutoSummarize(object length)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "AutoSummarize", NetOffice.WordApi.Range.LateBindingApiWrapperType, length);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="length">optional object length</param>
		/// <param name="mode">optional object mode</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range AutoSummarize(object length, object mode)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "AutoSummarize", NetOffice.WordApi.Range.LateBindingApiWrapperType, length, mode);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.RemoveNumbers"/> </remarks>
		/// <param name="numberType">optional object numberType</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void RemoveNumbers(object numberType)
		{
			 Factory.ExecuteMethod(this, "RemoveNumbers", numberType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.RemoveNumbers"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void RemoveNumbers()
		{
			 Factory.ExecuteMethod(this, "RemoveNumbers");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ConvertNumbersToText"/> </remarks>
		/// <param name="numberType">optional object numberType</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ConvertNumbersToText(object numberType)
		{
			 Factory.ExecuteMethod(this, "ConvertNumbersToText", numberType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ConvertNumbersToText"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ConvertNumbersToText()
		{
			 Factory.ExecuteMethod(this, "ConvertNumbersToText");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CountNumberedItems"/> </remarks>
		/// <param name="numberType">optional object numberType</param>
		/// <param name="level">optional object level</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 CountNumberedItems(object numberType, object level)
		{
			return Factory.ExecuteInt32MethodGet(this, "CountNumberedItems", numberType, level);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CountNumberedItems"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 CountNumberedItems()
		{
			return Factory.ExecuteInt32MethodGet(this, "CountNumberedItems");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CountNumberedItems"/> </remarks>
		/// <param name="numberType">optional object numberType</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 CountNumberedItems(object numberType)
		{
			return Factory.ExecuteInt32MethodGet(this, "CountNumberedItems", numberType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Post"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Post()
		{
			 Factory.ExecuteMethod(this, "Post");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ToggleFormsDesign"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ToggleFormsDesign()
		{
			 Factory.ExecuteMethod(this, "ToggleFormsDesign");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Compare"/> </remarks>
		/// <param name="name">string name</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Compare(string name)
		{
			 Factory.ExecuteMethod(this, "Compare", name);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Compare"/> </remarks>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		/// <param name="compareTarget">optional object compareTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		/// <param name="ignoreAllComparisonWarnings">optional object ignoreAllComparisonWarnings</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Compare(string name, object authorName, object compareTarget, object detectFormatChanges, object ignoreAllComparisonWarnings, object addToRecentFiles)
		{
			 Factory.ExecuteMethod(this, "Compare", new object[]{ name, authorName, compareTarget, detectFormatChanges, ignoreAllComparisonWarnings, addToRecentFiles });
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Compare"/> </remarks>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		/// <param name="compareTarget">optional object compareTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		/// <param name="ignoreAllComparisonWarnings">optional object ignoreAllComparisonWarnings</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="removePersonalInformation">optional object removePersonalInformation</param>
		/// <param name="removeDateAndTime">optional object removeDateAndTime</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void Compare(string name, object authorName, object compareTarget, object detectFormatChanges, object ignoreAllComparisonWarnings, object addToRecentFiles, object removePersonalInformation, object removeDateAndTime)
		{
			 Factory.ExecuteMethod(this, "Compare", new object[]{ name, authorName, compareTarget, detectFormatChanges, ignoreAllComparisonWarnings, addToRecentFiles, removePersonalInformation, removeDateAndTime });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Compare"/> </remarks>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Compare(string name, object authorName)
		{
			 Factory.ExecuteMethod(this, "Compare", name, authorName);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Compare"/> </remarks>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		/// <param name="compareTarget">optional object compareTarget</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Compare(string name, object authorName, object compareTarget)
		{
			 Factory.ExecuteMethod(this, "Compare", name, authorName, compareTarget);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Compare"/> </remarks>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		/// <param name="compareTarget">optional object compareTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Compare(string name, object authorName, object compareTarget, object detectFormatChanges)
		{
			 Factory.ExecuteMethod(this, "Compare", name, authorName, compareTarget, detectFormatChanges);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Compare"/> </remarks>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		/// <param name="compareTarget">optional object compareTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		/// <param name="ignoreAllComparisonWarnings">optional object ignoreAllComparisonWarnings</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Compare(string name, object authorName, object compareTarget, object detectFormatChanges, object ignoreAllComparisonWarnings)
		{
			 Factory.ExecuteMethod(this, "Compare", new object[]{ name, authorName, compareTarget, detectFormatChanges, ignoreAllComparisonWarnings });
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Compare"/> </remarks>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		/// <param name="compareTarget">optional object compareTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		/// <param name="ignoreAllComparisonWarnings">optional object ignoreAllComparisonWarnings</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="removePersonalInformation">optional object removePersonalInformation</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void Compare(string name, object authorName, object compareTarget, object detectFormatChanges, object ignoreAllComparisonWarnings, object addToRecentFiles, object removePersonalInformation)
		{
			 Factory.ExecuteMethod(this, "Compare", new object[]{ name, authorName, compareTarget, detectFormatChanges, ignoreAllComparisonWarnings, addToRecentFiles, removePersonalInformation });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void UpdateSummaryProperties()
		{
			 Factory.ExecuteMethod(this, "UpdateSummaryProperties");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.GetCrossReferenceItems"/> </remarks>
		/// <param name="referenceType">object referenceType</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object GetCrossReferenceItems(object referenceType)
		{
			return Factory.ExecuteVariantMethodGet(this, "GetCrossReferenceItems", referenceType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.AutoFormat"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void AutoFormat()
		{
			 Factory.ExecuteMethod(this, "AutoFormat");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ViewCode"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ViewCode()
		{
			 Factory.ExecuteMethod(this, "ViewCode");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ViewPropertyBrowser"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ViewPropertyBrowser()
		{
			 Factory.ExecuteMethod(this, "ViewPropertyBrowser");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ForwardMailer()
		{
			 Factory.ExecuteMethod(this, "ForwardMailer");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Reply()
		{
			 Factory.ExecuteMethod(this, "Reply");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ReplyAll()
		{
			 Factory.ExecuteMethod(this, "ReplyAll");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="priority">optional object priority</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SendMailer(object fileFormat, object priority)
		{
			 Factory.ExecuteMethod(this, "SendMailer", fileFormat, priority);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SendMailer()
		{
			 Factory.ExecuteMethod(this, "SendMailer");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileFormat">optional object fileFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SendMailer(object fileFormat)
		{
			 Factory.ExecuteMethod(this, "SendMailer", fileFormat);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.UndoClear"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void UndoClear()
		{
			 Factory.ExecuteMethod(this, "UndoClear");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PresentIt"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PresentIt()
		{
			 Factory.ExecuteMethod(this, "PresentIt");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SendFax"/> </remarks>
		/// <param name="address">string address</param>
		/// <param name="subject">optional object subject</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SendFax(string address, object subject)
		{
			 Factory.ExecuteMethod(this, "SendFax", address, subject);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SendFax"/> </remarks>
		/// <param name="address">string address</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SendFax(string address)
		{
			 Factory.ExecuteMethod(this, "SendFax", address);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Merge"/> </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Merge(string fileName)
		{
			 Factory.ExecuteMethod(this, "Merge", fileName);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Merge"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="mergeTarget">optional object mergeTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		/// <param name="useFormattingFrom">optional object useFormattingFrom</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Merge(string fileName, object mergeTarget, object detectFormatChanges, object useFormattingFrom, object addToRecentFiles)
		{
			 Factory.ExecuteMethod(this, "Merge", new object[]{ fileName, mergeTarget, detectFormatChanges, useFormattingFrom, addToRecentFiles });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Merge"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="mergeTarget">optional object mergeTarget</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Merge(string fileName, object mergeTarget)
		{
			 Factory.ExecuteMethod(this, "Merge", fileName, mergeTarget);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Merge"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="mergeTarget">optional object mergeTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Merge(string fileName, object mergeTarget, object detectFormatChanges)
		{
			 Factory.ExecuteMethod(this, "Merge", fileName, mergeTarget, detectFormatChanges);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Merge"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="mergeTarget">optional object mergeTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		/// <param name="useFormattingFrom">optional object useFormattingFrom</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Merge(string fileName, object mergeTarget, object detectFormatChanges, object useFormattingFrom)
		{
			 Factory.ExecuteMethod(this, "Merge", fileName, mergeTarget, detectFormatChanges, useFormattingFrom);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ClosePrintPreview"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ClosePrintPreview()
		{
			 Factory.ExecuteMethod(this, "ClosePrintPreview");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CheckConsistency"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckConsistency()
		{
			 Factory.ExecuteMethod(this, "CheckConsistency");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CreateLetterContent"/> </remarks>
		/// <param name="dateFormat">string dateFormat</param>
		/// <param name="includeHeaderFooter">bool includeHeaderFooter</param>
		/// <param name="pageDesign">string pageDesign</param>
		/// <param name="letterStyle">NetOffice.WordApi.Enums.WdLetterStyle letterStyle</param>
		/// <param name="letterhead">bool letterhead</param>
		/// <param name="letterheadLocation">NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation</param>
		/// <param name="letterheadSize">Single letterheadSize</param>
		/// <param name="recipientName">string recipientName</param>
		/// <param name="recipientAddress">string recipientAddress</param>
		/// <param name="salutation">string salutation</param>
		/// <param name="salutationType">NetOffice.WordApi.Enums.WdSalutationType salutationType</param>
		/// <param name="recipientReference">string recipientReference</param>
		/// <param name="mailingInstructions">string mailingInstructions</param>
		/// <param name="attentionLine">string attentionLine</param>
		/// <param name="subject">string subject</param>
		/// <param name="cCList">string cCList</param>
		/// <param name="returnAddress">string returnAddress</param>
		/// <param name="senderName">string senderName</param>
		/// <param name="closing">string closing</param>
		/// <param name="senderCompany">string senderCompany</param>
		/// <param name="senderJobTitle">string senderJobTitle</param>
		/// <param name="senderInitials">string senderInitials</param>
		/// <param name="enclosureNumber">Int32 enclosureNumber</param>
		/// <param name="infoBlock">optional object infoBlock</param>
		/// <param name="recipientCode">optional object recipientCode</param>
		/// <param name="recipientGender">optional object recipientGender</param>
		/// <param name="returnAddressShortForm">optional object returnAddressShortForm</param>
		/// <param name="senderCity">optional object senderCity</param>
		/// <param name="senderCode">optional object senderCode</param>
		/// <param name="senderGender">optional object senderGender</param>
		/// <param name="senderReference">optional object senderReference</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode, object recipientGender, object returnAddressShortForm, object senderCity, object senderCode, object senderGender, object senderReference)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.LetterContent>(this, "CreateLetterContent", NetOffice.WordApi.LetterContent.LateBindingApiWrapperType, new object[]{ dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cCList, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber, infoBlock, recipientCode, recipientGender, returnAddressShortForm, senderCity, senderCode, senderGender, senderReference });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CreateLetterContent"/> </remarks>
		/// <param name="dateFormat">string dateFormat</param>
		/// <param name="includeHeaderFooter">bool includeHeaderFooter</param>
		/// <param name="pageDesign">string pageDesign</param>
		/// <param name="letterStyle">NetOffice.WordApi.Enums.WdLetterStyle letterStyle</param>
		/// <param name="letterhead">bool letterhead</param>
		/// <param name="letterheadLocation">NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation</param>
		/// <param name="letterheadSize">Single letterheadSize</param>
		/// <param name="recipientName">string recipientName</param>
		/// <param name="recipientAddress">string recipientAddress</param>
		/// <param name="salutation">string salutation</param>
		/// <param name="salutationType">NetOffice.WordApi.Enums.WdSalutationType salutationType</param>
		/// <param name="recipientReference">string recipientReference</param>
		/// <param name="mailingInstructions">string mailingInstructions</param>
		/// <param name="attentionLine">string attentionLine</param>
		/// <param name="subject">string subject</param>
		/// <param name="cCList">string cCList</param>
		/// <param name="returnAddress">string returnAddress</param>
		/// <param name="senderName">string senderName</param>
		/// <param name="closing">string closing</param>
		/// <param name="senderCompany">string senderCompany</param>
		/// <param name="senderJobTitle">string senderJobTitle</param>
		/// <param name="senderInitials">string senderInitials</param>
		/// <param name="enclosureNumber">Int32 enclosureNumber</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.LetterContent>(this, "CreateLetterContent", NetOffice.WordApi.LetterContent.LateBindingApiWrapperType, new object[]{ dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cCList, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CreateLetterContent"/> </remarks>
		/// <param name="dateFormat">string dateFormat</param>
		/// <param name="includeHeaderFooter">bool includeHeaderFooter</param>
		/// <param name="pageDesign">string pageDesign</param>
		/// <param name="letterStyle">NetOffice.WordApi.Enums.WdLetterStyle letterStyle</param>
		/// <param name="letterhead">bool letterhead</param>
		/// <param name="letterheadLocation">NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation</param>
		/// <param name="letterheadSize">Single letterheadSize</param>
		/// <param name="recipientName">string recipientName</param>
		/// <param name="recipientAddress">string recipientAddress</param>
		/// <param name="salutation">string salutation</param>
		/// <param name="salutationType">NetOffice.WordApi.Enums.WdSalutationType salutationType</param>
		/// <param name="recipientReference">string recipientReference</param>
		/// <param name="mailingInstructions">string mailingInstructions</param>
		/// <param name="attentionLine">string attentionLine</param>
		/// <param name="subject">string subject</param>
		/// <param name="cCList">string cCList</param>
		/// <param name="returnAddress">string returnAddress</param>
		/// <param name="senderName">string senderName</param>
		/// <param name="closing">string closing</param>
		/// <param name="senderCompany">string senderCompany</param>
		/// <param name="senderJobTitle">string senderJobTitle</param>
		/// <param name="senderInitials">string senderInitials</param>
		/// <param name="enclosureNumber">Int32 enclosureNumber</param>
		/// <param name="infoBlock">optional object infoBlock</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.LetterContent>(this, "CreateLetterContent", NetOffice.WordApi.LetterContent.LateBindingApiWrapperType, new object[]{ dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cCList, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber, infoBlock });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CreateLetterContent"/> </remarks>
		/// <param name="dateFormat">string dateFormat</param>
		/// <param name="includeHeaderFooter">bool includeHeaderFooter</param>
		/// <param name="pageDesign">string pageDesign</param>
		/// <param name="letterStyle">NetOffice.WordApi.Enums.WdLetterStyle letterStyle</param>
		/// <param name="letterhead">bool letterhead</param>
		/// <param name="letterheadLocation">NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation</param>
		/// <param name="letterheadSize">Single letterheadSize</param>
		/// <param name="recipientName">string recipientName</param>
		/// <param name="recipientAddress">string recipientAddress</param>
		/// <param name="salutation">string salutation</param>
		/// <param name="salutationType">NetOffice.WordApi.Enums.WdSalutationType salutationType</param>
		/// <param name="recipientReference">string recipientReference</param>
		/// <param name="mailingInstructions">string mailingInstructions</param>
		/// <param name="attentionLine">string attentionLine</param>
		/// <param name="subject">string subject</param>
		/// <param name="cCList">string cCList</param>
		/// <param name="returnAddress">string returnAddress</param>
		/// <param name="senderName">string senderName</param>
		/// <param name="closing">string closing</param>
		/// <param name="senderCompany">string senderCompany</param>
		/// <param name="senderJobTitle">string senderJobTitle</param>
		/// <param name="senderInitials">string senderInitials</param>
		/// <param name="enclosureNumber">Int32 enclosureNumber</param>
		/// <param name="infoBlock">optional object infoBlock</param>
		/// <param name="recipientCode">optional object recipientCode</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.LetterContent>(this, "CreateLetterContent", NetOffice.WordApi.LetterContent.LateBindingApiWrapperType, new object[]{ dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cCList, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber, infoBlock, recipientCode });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CreateLetterContent"/> </remarks>
		/// <param name="dateFormat">string dateFormat</param>
		/// <param name="includeHeaderFooter">bool includeHeaderFooter</param>
		/// <param name="pageDesign">string pageDesign</param>
		/// <param name="letterStyle">NetOffice.WordApi.Enums.WdLetterStyle letterStyle</param>
		/// <param name="letterhead">bool letterhead</param>
		/// <param name="letterheadLocation">NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation</param>
		/// <param name="letterheadSize">Single letterheadSize</param>
		/// <param name="recipientName">string recipientName</param>
		/// <param name="recipientAddress">string recipientAddress</param>
		/// <param name="salutation">string salutation</param>
		/// <param name="salutationType">NetOffice.WordApi.Enums.WdSalutationType salutationType</param>
		/// <param name="recipientReference">string recipientReference</param>
		/// <param name="mailingInstructions">string mailingInstructions</param>
		/// <param name="attentionLine">string attentionLine</param>
		/// <param name="subject">string subject</param>
		/// <param name="cCList">string cCList</param>
		/// <param name="returnAddress">string returnAddress</param>
		/// <param name="senderName">string senderName</param>
		/// <param name="closing">string closing</param>
		/// <param name="senderCompany">string senderCompany</param>
		/// <param name="senderJobTitle">string senderJobTitle</param>
		/// <param name="senderInitials">string senderInitials</param>
		/// <param name="enclosureNumber">Int32 enclosureNumber</param>
		/// <param name="infoBlock">optional object infoBlock</param>
		/// <param name="recipientCode">optional object recipientCode</param>
		/// <param name="recipientGender">optional object recipientGender</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode, object recipientGender)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.LetterContent>(this, "CreateLetterContent", NetOffice.WordApi.LetterContent.LateBindingApiWrapperType, new object[]{ dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cCList, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber, infoBlock, recipientCode, recipientGender });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CreateLetterContent"/> </remarks>
		/// <param name="dateFormat">string dateFormat</param>
		/// <param name="includeHeaderFooter">bool includeHeaderFooter</param>
		/// <param name="pageDesign">string pageDesign</param>
		/// <param name="letterStyle">NetOffice.WordApi.Enums.WdLetterStyle letterStyle</param>
		/// <param name="letterhead">bool letterhead</param>
		/// <param name="letterheadLocation">NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation</param>
		/// <param name="letterheadSize">Single letterheadSize</param>
		/// <param name="recipientName">string recipientName</param>
		/// <param name="recipientAddress">string recipientAddress</param>
		/// <param name="salutation">string salutation</param>
		/// <param name="salutationType">NetOffice.WordApi.Enums.WdSalutationType salutationType</param>
		/// <param name="recipientReference">string recipientReference</param>
		/// <param name="mailingInstructions">string mailingInstructions</param>
		/// <param name="attentionLine">string attentionLine</param>
		/// <param name="subject">string subject</param>
		/// <param name="cCList">string cCList</param>
		/// <param name="returnAddress">string returnAddress</param>
		/// <param name="senderName">string senderName</param>
		/// <param name="closing">string closing</param>
		/// <param name="senderCompany">string senderCompany</param>
		/// <param name="senderJobTitle">string senderJobTitle</param>
		/// <param name="senderInitials">string senderInitials</param>
		/// <param name="enclosureNumber">Int32 enclosureNumber</param>
		/// <param name="infoBlock">optional object infoBlock</param>
		/// <param name="recipientCode">optional object recipientCode</param>
		/// <param name="recipientGender">optional object recipientGender</param>
		/// <param name="returnAddressShortForm">optional object returnAddressShortForm</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode, object recipientGender, object returnAddressShortForm)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.LetterContent>(this, "CreateLetterContent", NetOffice.WordApi.LetterContent.LateBindingApiWrapperType, new object[]{ dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cCList, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber, infoBlock, recipientCode, recipientGender, returnAddressShortForm });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CreateLetterContent"/> </remarks>
		/// <param name="dateFormat">string dateFormat</param>
		/// <param name="includeHeaderFooter">bool includeHeaderFooter</param>
		/// <param name="pageDesign">string pageDesign</param>
		/// <param name="letterStyle">NetOffice.WordApi.Enums.WdLetterStyle letterStyle</param>
		/// <param name="letterhead">bool letterhead</param>
		/// <param name="letterheadLocation">NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation</param>
		/// <param name="letterheadSize">Single letterheadSize</param>
		/// <param name="recipientName">string recipientName</param>
		/// <param name="recipientAddress">string recipientAddress</param>
		/// <param name="salutation">string salutation</param>
		/// <param name="salutationType">NetOffice.WordApi.Enums.WdSalutationType salutationType</param>
		/// <param name="recipientReference">string recipientReference</param>
		/// <param name="mailingInstructions">string mailingInstructions</param>
		/// <param name="attentionLine">string attentionLine</param>
		/// <param name="subject">string subject</param>
		/// <param name="cCList">string cCList</param>
		/// <param name="returnAddress">string returnAddress</param>
		/// <param name="senderName">string senderName</param>
		/// <param name="closing">string closing</param>
		/// <param name="senderCompany">string senderCompany</param>
		/// <param name="senderJobTitle">string senderJobTitle</param>
		/// <param name="senderInitials">string senderInitials</param>
		/// <param name="enclosureNumber">Int32 enclosureNumber</param>
		/// <param name="infoBlock">optional object infoBlock</param>
		/// <param name="recipientCode">optional object recipientCode</param>
		/// <param name="recipientGender">optional object recipientGender</param>
		/// <param name="returnAddressShortForm">optional object returnAddressShortForm</param>
		/// <param name="senderCity">optional object senderCity</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode, object recipientGender, object returnAddressShortForm, object senderCity)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.LetterContent>(this, "CreateLetterContent", NetOffice.WordApi.LetterContent.LateBindingApiWrapperType, new object[]{ dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cCList, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber, infoBlock, recipientCode, recipientGender, returnAddressShortForm, senderCity });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CreateLetterContent"/> </remarks>
		/// <param name="dateFormat">string dateFormat</param>
		/// <param name="includeHeaderFooter">bool includeHeaderFooter</param>
		/// <param name="pageDesign">string pageDesign</param>
		/// <param name="letterStyle">NetOffice.WordApi.Enums.WdLetterStyle letterStyle</param>
		/// <param name="letterhead">bool letterhead</param>
		/// <param name="letterheadLocation">NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation</param>
		/// <param name="letterheadSize">Single letterheadSize</param>
		/// <param name="recipientName">string recipientName</param>
		/// <param name="recipientAddress">string recipientAddress</param>
		/// <param name="salutation">string salutation</param>
		/// <param name="salutationType">NetOffice.WordApi.Enums.WdSalutationType salutationType</param>
		/// <param name="recipientReference">string recipientReference</param>
		/// <param name="mailingInstructions">string mailingInstructions</param>
		/// <param name="attentionLine">string attentionLine</param>
		/// <param name="subject">string subject</param>
		/// <param name="cCList">string cCList</param>
		/// <param name="returnAddress">string returnAddress</param>
		/// <param name="senderName">string senderName</param>
		/// <param name="closing">string closing</param>
		/// <param name="senderCompany">string senderCompany</param>
		/// <param name="senderJobTitle">string senderJobTitle</param>
		/// <param name="senderInitials">string senderInitials</param>
		/// <param name="enclosureNumber">Int32 enclosureNumber</param>
		/// <param name="infoBlock">optional object infoBlock</param>
		/// <param name="recipientCode">optional object recipientCode</param>
		/// <param name="recipientGender">optional object recipientGender</param>
		/// <param name="returnAddressShortForm">optional object returnAddressShortForm</param>
		/// <param name="senderCity">optional object senderCity</param>
		/// <param name="senderCode">optional object senderCode</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode, object recipientGender, object returnAddressShortForm, object senderCity, object senderCode)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.LetterContent>(this, "CreateLetterContent", NetOffice.WordApi.LetterContent.LateBindingApiWrapperType, new object[]{ dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cCList, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber, infoBlock, recipientCode, recipientGender, returnAddressShortForm, senderCity, senderCode });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CreateLetterContent"/> </remarks>
		/// <param name="dateFormat">string dateFormat</param>
		/// <param name="includeHeaderFooter">bool includeHeaderFooter</param>
		/// <param name="pageDesign">string pageDesign</param>
		/// <param name="letterStyle">NetOffice.WordApi.Enums.WdLetterStyle letterStyle</param>
		/// <param name="letterhead">bool letterhead</param>
		/// <param name="letterheadLocation">NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation</param>
		/// <param name="letterheadSize">Single letterheadSize</param>
		/// <param name="recipientName">string recipientName</param>
		/// <param name="recipientAddress">string recipientAddress</param>
		/// <param name="salutation">string salutation</param>
		/// <param name="salutationType">NetOffice.WordApi.Enums.WdSalutationType salutationType</param>
		/// <param name="recipientReference">string recipientReference</param>
		/// <param name="mailingInstructions">string mailingInstructions</param>
		/// <param name="attentionLine">string attentionLine</param>
		/// <param name="subject">string subject</param>
		/// <param name="cCList">string cCList</param>
		/// <param name="returnAddress">string returnAddress</param>
		/// <param name="senderName">string senderName</param>
		/// <param name="closing">string closing</param>
		/// <param name="senderCompany">string senderCompany</param>
		/// <param name="senderJobTitle">string senderJobTitle</param>
		/// <param name="senderInitials">string senderInitials</param>
		/// <param name="enclosureNumber">Int32 enclosureNumber</param>
		/// <param name="infoBlock">optional object infoBlock</param>
		/// <param name="recipientCode">optional object recipientCode</param>
		/// <param name="recipientGender">optional object recipientGender</param>
		/// <param name="returnAddressShortForm">optional object returnAddressShortForm</param>
		/// <param name="senderCity">optional object senderCity</param>
		/// <param name="senderCode">optional object senderCode</param>
		/// <param name="senderGender">optional object senderGender</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode, object recipientGender, object returnAddressShortForm, object senderCity, object senderCode, object senderGender)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.LetterContent>(this, "CreateLetterContent", NetOffice.WordApi.LetterContent.LateBindingApiWrapperType, new object[]{ dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cCList, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber, infoBlock, recipientCode, recipientGender, returnAddressShortForm, senderCity, senderCode, senderGender });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.AcceptAllRevisions"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void AcceptAllRevisions()
		{
			 Factory.ExecuteMethod(this, "AcceptAllRevisions");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.RejectAllRevisions"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void RejectAllRevisions()
		{
			 Factory.ExecuteMethod(this, "RejectAllRevisions");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.DetectLanguage"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void DetectLanguage()
		{
			 Factory.ExecuteMethod(this, "DetectLanguage");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ApplyTheme"/> </remarks>
		/// <param name="name">string name</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ApplyTheme(string name)
		{
			 Factory.ExecuteMethod(this, "ApplyTheme", name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.RemoveTheme"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void RemoveTheme()
		{
			 Factory.ExecuteMethod(this, "RemoveTheme");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.WebPagePreview"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void WebPagePreview()
		{
			 Factory.ExecuteMethod(this, "WebPagePreview");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ReloadAs"/> </remarks>
		/// <param name="encoding">NetOffice.OfficeApi.Enums.MsoEncoding encoding</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ReloadAs(NetOffice.OfficeApi.Enums.MsoEncoding encoding)
		{
			 Factory.ExecuteMethod(this, "ReloadAs", encoding);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PrintOut"/> </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		/// <param name="printZoomRow">optional object printZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object printZoomPaperWidth</param>
		/// <param name="printZoomPaperHeight">optional object printZoomPaperHeight</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth, object printZoomPaperHeight)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth, printZoomPaperHeight });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PrintOut"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut()
		{
			 Factory.ExecuteMethod(this, "PrintOut");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PrintOut"/> </remarks>
		/// <param name="background">optional object background</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background)
		{
			 Factory.ExecuteMethod(this, "PrintOut", background);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PrintOut"/> </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append)
		{
			 Factory.ExecuteMethod(this, "PrintOut", background, append);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PrintOut"/> </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range)
		{
			 Factory.ExecuteMethod(this, "PrintOut", background, append, range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PrintOut"/> </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName)
		{
			 Factory.ExecuteMethod(this, "PrintOut", background, append, range, outputFileName);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PrintOut"/> </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PrintOut"/> </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PrintOut"/> </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PrintOut"/> </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PrintOut"/> </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PrintOut"/> </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PrintOut"/> </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PrintOut"/> </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PrintOut"/> </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PrintOut"/> </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PrintOut"/> </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PrintOut"/> </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		/// <param name="printZoomRow">optional object printZoomRow</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.PrintOut"/> </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		/// <param name="printZoomRow">optional object printZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object printZoomPaperWidth</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="s">string s</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void sblt(string s)
		{
			 Factory.ExecuteMethod(this, "sblt", s);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter)
		{
			 Factory.ExecuteMethod(this, "SaveAs2000", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SaveAs2000()
		{
			 Factory.ExecuteMethod(this, "SaveAs2000");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SaveAs2000(object fileName)
		{
			 Factory.ExecuteMethod(this, "SaveAs2000", fileName);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SaveAs2000(object fileName, object fileFormat)
		{
			 Factory.ExecuteMethod(this, "SaveAs2000", fileName, fileFormat);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SaveAs2000(object fileName, object fileFormat, object lockComments)
		{
			 Factory.ExecuteMethod(this, "SaveAs2000", fileName, fileFormat, lockComments);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SaveAs2000(object fileName, object fileFormat, object lockComments, object password)
		{
			 Factory.ExecuteMethod(this, "SaveAs2000", fileName, fileFormat, lockComments, password);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles)
		{
			 Factory.ExecuteMethod(this, "SaveAs2000", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword)
		{
			 Factory.ExecuteMethod(this, "SaveAs2000", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended)
		{
			 Factory.ExecuteMethod(this, "SaveAs2000", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts)
		{
			 Factory.ExecuteMethod(this, "SaveAs2000", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat)
		{
			 Factory.ExecuteMethod(this, "SaveAs2000", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData)
		{
			 Factory.ExecuteMethod(this, "SaveAs2000", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Compare2000(string name)
		{
			 Factory.ExecuteMethod(this, "Compare2000", name);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Merge2000(string fileName)
		{
			 Factory.ExecuteMethod(this, "Merge2000", fileName);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		/// <param name="printZoomRow">optional object printZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object printZoomPaperWidth</param>
		/// <param name="printZoomPaperHeight">optional object printZoomPaperHeight</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth, object printZoomPaperHeight)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth, printZoomPaperHeight });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000()
		{
			 Factory.ExecuteMethod(this, "PrintOut2000");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", background);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", background, append);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", background, append, range);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", background, append, range, outputFileName);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		/// <param name="printZoomRow">optional object printZoomRow</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		/// <param name="printZoomRow">optional object printZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object printZoomPaperWidth</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ConvertVietDoc"/> </remarks>
		/// <param name="codePageOrigin">Int32 codePageOrigin</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void ConvertVietDoc(Int32 codePageOrigin)
		{
			 Factory.ExecuteMethod(this, "ConvertVietDoc", codePageOrigin);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CheckIn"/> </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		/// <param name="makePublic">optional bool MakePublic = false</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void CheckIn(object saveChanges, object comments, object makePublic)
		{
			 Factory.ExecuteMethod(this, "CheckIn", saveChanges, comments, makePublic);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CheckIn"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void CheckIn()
		{
			 Factory.ExecuteMethod(this, "CheckIn");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CheckIn"/> </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void CheckIn(object saveChanges)
		{
			 Factory.ExecuteMethod(this, "CheckIn", saveChanges);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CheckIn"/> </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void CheckIn(object saveChanges, object comments)
		{
			 Factory.ExecuteMethod(this, "CheckIn", saveChanges, comments);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CanCheckin"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool CanCheckin()
		{
			return Factory.ExecuteBoolMethodGet(this, "CanCheckin");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SendForReview"/> </remarks>
		/// <param name="recipients">optional object recipients</param>
		/// <param name="subject">optional object subject</param>
		/// <param name="showMessage">optional object showMessage</param>
		/// <param name="includeAttachment">optional object includeAttachment</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SendForReview(object recipients, object subject, object showMessage, object includeAttachment)
		{
			 Factory.ExecuteMethod(this, "SendForReview", recipients, subject, showMessage, includeAttachment);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SendForReview"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SendForReview()
		{
			 Factory.ExecuteMethod(this, "SendForReview");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SendForReview"/> </remarks>
		/// <param name="recipients">optional object recipients</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SendForReview(object recipients)
		{
			 Factory.ExecuteMethod(this, "SendForReview", recipients);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SendForReview"/> </remarks>
		/// <param name="recipients">optional object recipients</param>
		/// <param name="subject">optional object subject</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SendForReview(object recipients, object subject)
		{
			 Factory.ExecuteMethod(this, "SendForReview", recipients, subject);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SendForReview"/> </remarks>
		/// <param name="recipients">optional object recipients</param>
		/// <param name="subject">optional object subject</param>
		/// <param name="showMessage">optional object showMessage</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SendForReview(object recipients, object subject, object showMessage)
		{
			 Factory.ExecuteMethod(this, "SendForReview", recipients, subject, showMessage);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ReplyWithChanges"/> </remarks>
		/// <param name="showMessage">optional object showMessage</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void ReplyWithChanges(object showMessage)
		{
			 Factory.ExecuteMethod(this, "ReplyWithChanges", showMessage);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ReplyWithChanges"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void ReplyWithChanges()
		{
			 Factory.ExecuteMethod(this, "ReplyWithChanges");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.EndReview"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void EndReview()
		{
			 Factory.ExecuteMethod(this, "EndReview");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SetPasswordEncryptionOptions"/> </remarks>
		/// <param name="passwordEncryptionProvider">string passwordEncryptionProvider</param>
		/// <param name="passwordEncryptionAlgorithm">string passwordEncryptionAlgorithm</param>
		/// <param name="passwordEncryptionKeyLength">Int32 passwordEncryptionKeyLength</param>
		/// <param name="passwordEncryptionFileProperties">optional object passwordEncryptionFileProperties</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SetPasswordEncryptionOptions(string passwordEncryptionProvider, string passwordEncryptionAlgorithm, Int32 passwordEncryptionKeyLength, object passwordEncryptionFileProperties)
		{
			 Factory.ExecuteMethod(this, "SetPasswordEncryptionOptions", passwordEncryptionProvider, passwordEncryptionAlgorithm, passwordEncryptionKeyLength, passwordEncryptionFileProperties);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SetPasswordEncryptionOptions"/> </remarks>
		/// <param name="passwordEncryptionProvider">string passwordEncryptionProvider</param>
		/// <param name="passwordEncryptionAlgorithm">string passwordEncryptionAlgorithm</param>
		/// <param name="passwordEncryptionKeyLength">Int32 passwordEncryptionKeyLength</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SetPasswordEncryptionOptions(string passwordEncryptionProvider, string passwordEncryptionAlgorithm, Int32 passwordEncryptionKeyLength)
		{
			 Factory.ExecuteMethod(this, "SetPasswordEncryptionOptions", passwordEncryptionProvider, passwordEncryptionAlgorithm, passwordEncryptionKeyLength);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void RecheckSmartTags()
		{
			 Factory.ExecuteMethod(this, "RecheckSmartTags");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void RemoveSmartTags()
		{
			 Factory.ExecuteMethod(this, "RemoveSmartTags");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SetDefaultTableStyle"/> </remarks>
		/// <param name="style">object style</param>
		/// <param name="setInTemplate">bool setInTemplate</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SetDefaultTableStyle(object style, bool setInTemplate)
		{
			 Factory.ExecuteMethod(this, "SetDefaultTableStyle", style, setInTemplate);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.DeleteAllComments"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void DeleteAllComments()
		{
			 Factory.ExecuteMethod(this, "DeleteAllComments");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.AcceptAllRevisionsShown"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void AcceptAllRevisionsShown()
		{
			 Factory.ExecuteMethod(this, "AcceptAllRevisionsShown");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.RejectAllRevisionsShown"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void RejectAllRevisionsShown()
		{
			 Factory.ExecuteMethod(this, "RejectAllRevisionsShown");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.DeleteAllCommentsShown"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void DeleteAllCommentsShown()
		{
			 Factory.ExecuteMethod(this, "DeleteAllCommentsShown");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ResetFormFields"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void ResetFormFields()
		{
			 Factory.ExecuteMethod(this, "ResetFormFields");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void CheckNewSmartTags()
		{
			 Factory.ExecuteMethod(this, "CheckNewSmartTags");
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType type</param>
		/// <param name="noReset">optional object noReset</param>
		/// <param name="password">optional object password</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void Protect2002(NetOffice.WordApi.Enums.WdProtectionType type, object noReset, object password)
		{
			 Factory.ExecuteMethod(this, "Protect2002", type, noReset, password);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType type</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void Protect2002(NetOffice.WordApi.Enums.WdProtectionType type)
		{
			 Factory.ExecuteMethod(this, "Protect2002", type);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType type</param>
		/// <param name="noReset">optional object noReset</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void Protect2002(NetOffice.WordApi.Enums.WdProtectionType type, object noReset)
		{
			 Factory.ExecuteMethod(this, "Protect2002", type, noReset);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		/// <param name="compareTarget">optional object compareTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		/// <param name="ignoreAllComparisonWarnings">optional object ignoreAllComparisonWarnings</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void Compare2002(string name, object authorName, object compareTarget, object detectFormatChanges, object ignoreAllComparisonWarnings, object addToRecentFiles)
		{
			 Factory.ExecuteMethod(this, "Compare2002", new object[]{ name, authorName, compareTarget, detectFormatChanges, ignoreAllComparisonWarnings, addToRecentFiles });
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void Compare2002(string name)
		{
			 Factory.ExecuteMethod(this, "Compare2002", name);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void Compare2002(string name, object authorName)
		{
			 Factory.ExecuteMethod(this, "Compare2002", name, authorName);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		/// <param name="compareTarget">optional object compareTarget</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void Compare2002(string name, object authorName, object compareTarget)
		{
			 Factory.ExecuteMethod(this, "Compare2002", name, authorName, compareTarget);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		/// <param name="compareTarget">optional object compareTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void Compare2002(string name, object authorName, object compareTarget, object detectFormatChanges)
		{
			 Factory.ExecuteMethod(this, "Compare2002", name, authorName, compareTarget, detectFormatChanges);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		/// <param name="compareTarget">optional object compareTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		/// <param name="ignoreAllComparisonWarnings">optional object ignoreAllComparisonWarnings</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void Compare2002(string name, object authorName, object compareTarget, object detectFormatChanges, object ignoreAllComparisonWarnings)
		{
			 Factory.ExecuteMethod(this, "Compare2002", new object[]{ name, authorName, compareTarget, detectFormatChanges, ignoreAllComparisonWarnings });
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SendFaxOverInternet"/> </remarks>
		/// <param name="recipients">optional object recipients</param>
		/// <param name="subject">optional object subject</param>
		/// <param name="showMessage">optional object showMessage</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void SendFaxOverInternet(object recipients, object subject, object showMessage)
		{
			 Factory.ExecuteMethod(this, "SendFaxOverInternet", recipients, subject, showMessage);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SendFaxOverInternet"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void SendFaxOverInternet()
		{
			 Factory.ExecuteMethod(this, "SendFaxOverInternet");
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SendFaxOverInternet"/> </remarks>
		/// <param name="recipients">optional object recipients</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void SendFaxOverInternet(object recipients)
		{
			 Factory.ExecuteMethod(this, "SendFaxOverInternet", recipients);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SendFaxOverInternet"/> </remarks>
		/// <param name="recipients">optional object recipients</param>
		/// <param name="subject">optional object subject</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void SendFaxOverInternet(object recipients, object subject)
		{
			 Factory.ExecuteMethod(this, "SendFaxOverInternet", recipients, subject);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.TransformDocument"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="dataOnly">optional bool DataOnly = true</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void TransformDocument(string path, object dataOnly)
		{
			 Factory.ExecuteMethod(this, "TransformDocument", path, dataOnly);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.TransformDocument"/> </remarks>
		/// <param name="path">string path</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void TransformDocument(string path)
		{
			 Factory.ExecuteMethod(this, "TransformDocument", path);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SelectAllEditableRanges"/> </remarks>
		/// <param name="editorID">optional object editorID</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void SelectAllEditableRanges(object editorID)
		{
			 Factory.ExecuteMethod(this, "SelectAllEditableRanges", editorID);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SelectAllEditableRanges"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void SelectAllEditableRanges()
		{
			 Factory.ExecuteMethod(this, "SelectAllEditableRanges");
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.DeleteAllEditableRanges"/> </remarks>
		/// <param name="editorID">optional object editorID</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void DeleteAllEditableRanges(object editorID)
		{
			 Factory.ExecuteMethod(this, "DeleteAllEditableRanges", editorID);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.DeleteAllEditableRanges"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void DeleteAllEditableRanges()
		{
			 Factory.ExecuteMethod(this, "DeleteAllEditableRanges");
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.DeleteAllInkAnnotations"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void DeleteAllInkAnnotations()
		{
			 Factory.ExecuteMethod(this, "DeleteAllInkAnnotations");
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="richFormat">bool richFormat</param>
		/// <param name="url">string url</param>
		/// <param name="title">string title</param>
		/// <param name="description">string description</param>
		/// <param name="iD">string iD</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void AddDocumentWorkspaceHeader(bool richFormat, string url, string title, string description, string iD)
		{
			 Factory.ExecuteMethod(this, "AddDocumentWorkspaceHeader", new object[]{ richFormat, url, title, description, iD });
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="iD">string iD</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void RemoveDocumentWorkspaceHeader(string iD)
		{
			 Factory.ExecuteMethod(this, "RemoveDocumentWorkspaceHeader", iD);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.RemoveLockedStyles"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void RemoveLockedStyles()
		{
			 Factory.ExecuteMethod(this, "RemoveLockedStyles");
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SelectSingleNode"/> </remarks>
		/// <param name="xPath">string xPath</param>
		/// <param name="prefixMapping">optional string PrefixMapping = </param>
		/// <param name="fastSearchSkippingTextNodes">optional bool FastSearchSkippingTextNodes = true</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNode SelectSingleNode(string xPath, object prefixMapping, object fastSearchSkippingTextNodes)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLNode>(this, "SelectSingleNode", NetOffice.WordApi.XMLNode.LateBindingApiWrapperType, xPath, prefixMapping, fastSearchSkippingTextNodes);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SelectSingleNode"/> </remarks>
		/// <param name="xPath">string xPath</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNode SelectSingleNode(string xPath)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLNode>(this, "SelectSingleNode", NetOffice.WordApi.XMLNode.LateBindingApiWrapperType, xPath);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SelectSingleNode"/> </remarks>
		/// <param name="xPath">string xPath</param>
		/// <param name="prefixMapping">optional string PrefixMapping = </param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNode SelectSingleNode(string xPath, object prefixMapping)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLNode>(this, "SelectSingleNode", NetOffice.WordApi.XMLNode.LateBindingApiWrapperType, xPath, prefixMapping);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SelectNodes"/> </remarks>
		/// <param name="xPath">string xPath</param>
		/// <param name="prefixMapping">optional string PrefixMapping = </param>
		/// <param name="fastSearchSkippingTextNodes">optional bool FastSearchSkippingTextNodes = true</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNodes SelectNodes(string xPath, object prefixMapping, object fastSearchSkippingTextNodes)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLNodes>(this, "SelectNodes", NetOffice.WordApi.XMLNodes.LateBindingApiWrapperType, xPath, prefixMapping, fastSearchSkippingTextNodes);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SelectNodes"/> </remarks>
		/// <param name="xPath">string xPath</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNodes SelectNodes(string xPath)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLNodes>(this, "SelectNodes", NetOffice.WordApi.XMLNodes.LateBindingApiWrapperType, xPath);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SelectNodes"/> </remarks>
		/// <param name="xPath">string xPath</param>
		/// <param name="prefixMapping">optional string PrefixMapping = </param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNodes SelectNodes(string xPath, object prefixMapping)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLNodes>(this, "SelectNodes", NetOffice.WordApi.XMLNodes.LateBindingApiWrapperType, xPath, prefixMapping);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.RemoveDocumentInformation"/> </remarks>
		/// <param name="removeDocInfoType">NetOffice.WordApi.Enums.WdRemoveDocInfoType removeDocInfoType</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public void RemoveDocumentInformation(NetOffice.WordApi.Enums.WdRemoveDocInfoType removeDocInfoType)
		{
			 Factory.ExecuteMethod(this, "RemoveDocumentInformation", removeDocInfoType);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CheckInWithVersion"/> </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		/// <param name="makePublic">optional bool MakePublic = false</param>
		/// <param name="versionType">optional object versionType</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges, object comments, object makePublic, object versionType)
		{
			 Factory.ExecuteMethod(this, "CheckInWithVersion", saveChanges, comments, makePublic, versionType);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CheckInWithVersion"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void CheckInWithVersion()
		{
			 Factory.ExecuteMethod(this, "CheckInWithVersion");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CheckInWithVersion"/> </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges)
		{
			 Factory.ExecuteMethod(this, "CheckInWithVersion", saveChanges);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CheckInWithVersion"/> </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges, object comments)
		{
			 Factory.ExecuteMethod(this, "CheckInWithVersion", saveChanges, comments);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.CheckInWithVersion"/> </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		/// <param name="makePublic">optional bool MakePublic = false</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges, object comments, object makePublic)
		{
			 Factory.ExecuteMethod(this, "CheckInWithVersion", saveChanges, comments, makePublic);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 12,14,15,16)]
		public void Dummy2()
		{
			 Factory.ExecuteMethod(this, "Dummy2");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.LockServerFile"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void LockServerFile()
		{
			 Factory.ExecuteMethod(this, "LockServerFile");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.GetWorkflowTasks"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.OfficeApi.WorkflowTasks GetWorkflowTasks()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.WorkflowTasks>(this, "GetWorkflowTasks", NetOffice.OfficeApi.WorkflowTasks.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.GetWorkflowTemplates"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.OfficeApi.WorkflowTemplates GetWorkflowTemplates()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.WorkflowTemplates>(this, "GetWorkflowTemplates", NetOffice.OfficeApi.WorkflowTemplates.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 12,14,15,16)]
		public void Dummy4()
		{
			 Factory.ExecuteMethod(this, "Dummy4");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <param name="skipIfAbsent">bool skipIfAbsent</param>
		/// <param name="url">string url</param>
		/// <param name="title">string title</param>
		/// <param name="description">string description</param>
		/// <param name="iD">string iD</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 12,14,15,16)]
		public void AddMeetingWorkspaceHeader(bool skipIfAbsent, string url, string title, string description, string iD)
		{
			 Factory.ExecuteMethod(this, "AddMeetingWorkspaceHeader", new object[]{ skipIfAbsent, url, title, description, iD });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SaveAsQuickStyleSet"/> </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public void SaveAsQuickStyleSet(string fileName)
		{
			 Factory.ExecuteMethod(this, "SaveAsQuickStyleSet", fileName);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public void ApplyQuickStyleSet(string name)
		{
			 Factory.ExecuteMethod(this, "ApplyQuickStyleSet", name);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ApplyDocumentTheme"/> </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public void ApplyDocumentTheme(string fileName)
		{
			 Factory.ExecuteMethod(this, "ApplyDocumentTheme", fileName);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SelectLinkedControls"/> </remarks>
		/// <param name="node">NetOffice.OfficeApi.CustomXMLNode node</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.ContentControls SelectLinkedControls(NetOffice.OfficeApi.CustomXMLNode node)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.ContentControls>(this, "SelectLinkedControls", NetOffice.WordApi.ContentControls.LateBindingApiWrapperType, node);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SelectUnlinkedControls"/> </remarks>
		/// <param name="stream">optional NetOffice.OfficeApi.CustomXMLPart Stream = 0</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.ContentControls SelectUnlinkedControls(object stream)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.ContentControls>(this, "SelectUnlinkedControls", NetOffice.WordApi.ContentControls.LateBindingApiWrapperType, stream);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SelectUnlinkedControls"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.ContentControls SelectUnlinkedControls()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.ContentControls>(this, "SelectUnlinkedControls", NetOffice.WordApi.ContentControls.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SelectContentControlsByTitle"/> </remarks>
		/// <param name="title">string title</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.ContentControls SelectContentControlsByTitle(string title)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.ContentControls>(this, "SelectContentControlsByTitle", NetOffice.WordApi.ContentControls.LateBindingApiWrapperType, title);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ExportAsFixedFormat"/> </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		/// <param name="to">optional Int32 To = 1</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		/// <param name="keepIRM">optional bool KeepIRM = true</param>
		/// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		/// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
		/// <param name="fixedFormatExtClassPtr">optional object fixedFormatExtClassPtr</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags, object bitmapMissingFonts, object useISO19005_1, object fixedFormatExtClassPtr)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, range, from, to, item, includeDocProps, keepIRM, createBookmarks, docStructureTags, bitmapMissingFonts, useISO19005_1, fixedFormatExtClassPtr });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ExportAsFixedFormat"/> </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", outputFileName, exportFormat);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ExportAsFixedFormat"/> </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", outputFileName, exportFormat, openAfterExport);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ExportAsFixedFormat"/> </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", outputFileName, exportFormat, openAfterExport, optimizeFor);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ExportAsFixedFormat"/> </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, range });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ExportAsFixedFormat"/> </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, range, from });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ExportAsFixedFormat"/> </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		/// <param name="to">optional Int32 To = 1</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, range, from, to });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ExportAsFixedFormat"/> </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		/// <param name="to">optional Int32 To = 1</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, range, from, to, item });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ExportAsFixedFormat"/> </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		/// <param name="to">optional Int32 To = 1</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, range, from, to, item, includeDocProps });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ExportAsFixedFormat"/> </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		/// <param name="to">optional Int32 To = 1</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		/// <param name="keepIRM">optional bool KeepIRM = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps, object keepIRM)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, range, from, to, item, includeDocProps, keepIRM });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ExportAsFixedFormat"/> </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		/// <param name="to">optional Int32 To = 1</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		/// <param name="keepIRM">optional bool KeepIRM = true</param>
		/// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps, object keepIRM, object createBookmarks)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, range, from, to, item, includeDocProps, keepIRM, createBookmarks });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ExportAsFixedFormat"/> </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		/// <param name="to">optional Int32 To = 1</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		/// <param name="keepIRM">optional bool KeepIRM = true</param>
		/// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, range, from, to, item, includeDocProps, keepIRM, createBookmarks, docStructureTags });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ExportAsFixedFormat"/> </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		/// <param name="to">optional Int32 To = 1</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		/// <param name="keepIRM">optional bool KeepIRM = true</param>
		/// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags, object bitmapMissingFonts)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, range, from, to, item, includeDocProps, keepIRM, createBookmarks, docStructureTags, bitmapMissingFonts });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ExportAsFixedFormat"/> </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		/// <param name="to">optional Int32 To = 1</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		/// <param name="keepIRM">optional bool KeepIRM = true</param>
		/// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		/// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags, object bitmapMissingFonts, object useISO19005_1)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, range, from, to, item, includeDocProps, keepIRM, createBookmarks, docStructureTags, bitmapMissingFonts, useISO19005_1 });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.FreezeLayout"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void FreezeLayout()
		{
			 Factory.ExecuteMethod(this, "FreezeLayout");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 12,14,15,16)]
		public void UnfreezeLayout()
		{
			 Factory.ExecuteMethod(this, "UnfreezeLayout");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.DowngradeDocument"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void DowngradeDocument()
		{
			 Factory.ExecuteMethod(this, "DowngradeDocument");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.Convert"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void Convert()
		{
			 Factory.ExecuteMethod(this, "Convert");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.SelectContentControlsByTag"/> </remarks>
		/// <param name="tag">string tag</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.ContentControls SelectContentControlsByTag(string tag)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.ContentControls>(this, "SelectContentControlsByTag", NetOffice.WordApi.ContentControls.LateBindingApiWrapperType, tag);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ConvertAutoHyphens"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public void ConvertAutoHyphens()
		{
			 Factory.ExecuteMethod(this, "ConvertAutoHyphens");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Document.ApplyQuickStyleSet2"/> </remarks>
		/// <param name="style">object style</param>
		[SupportByVersion("Word", 14,15,16)]
		public void ApplyQuickStyleSet2(object style)
		{
			 Factory.ExecuteMethod(this, "ApplyQuickStyleSet2", style);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.SaveAs2"/> </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object allowSubstitutions</param>
		/// <param name="lineEnding">optional object lineEnding</param>
		/// <param name="addBiDiMarks">optional object addBiDiMarks</param>
		/// <param name="compatibilityMode">optional object compatibilityMode</param>
		[SupportByVersion("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding, object addBiDiMarks, object compatibilityMode)
		{
			 Factory.ExecuteMethod(this, "SaveAs2", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions, lineEnding, addBiDiMarks, compatibilityMode });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.SaveAs2"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void SaveAs2()
		{
			 Factory.ExecuteMethod(this, "SaveAs2");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.SaveAs2"/> </remarks>
		/// <param name="fileName">optional object fileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void SaveAs2(object fileName)
		{
			 Factory.ExecuteMethod(this, "SaveAs2", fileName);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.SaveAs2"/> </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat)
		{
			 Factory.ExecuteMethod(this, "SaveAs2", fileName, fileFormat);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.SaveAs2"/> </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments)
		{
			 Factory.ExecuteMethod(this, "SaveAs2", fileName, fileFormat, lockComments);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.SaveAs2"/> </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password)
		{
			 Factory.ExecuteMethod(this, "SaveAs2", fileName, fileFormat, lockComments, password);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.SaveAs2"/> </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles)
		{
			 Factory.ExecuteMethod(this, "SaveAs2", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.SaveAs2"/> </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword)
		{
			 Factory.ExecuteMethod(this, "SaveAs2", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.SaveAs2"/> </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended)
		{
			 Factory.ExecuteMethod(this, "SaveAs2", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.SaveAs2"/> </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts)
		{
			 Factory.ExecuteMethod(this, "SaveAs2", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.SaveAs2"/> </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat)
		{
			 Factory.ExecuteMethod(this, "SaveAs2", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.SaveAs2"/> </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData)
		{
			 Factory.ExecuteMethod(this, "SaveAs2", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.SaveAs2"/> </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter)
		{
			 Factory.ExecuteMethod(this, "SaveAs2", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.SaveAs2"/> </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding)
		{
			 Factory.ExecuteMethod(this, "SaveAs2", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.SaveAs2"/> </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks)
		{
			 Factory.ExecuteMethod(this, "SaveAs2", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.SaveAs2"/> </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object allowSubstitutions</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions)
		{
			 Factory.ExecuteMethod(this, "SaveAs2", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.SaveAs2"/> </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object allowSubstitutions</param>
		/// <param name="lineEnding">optional object lineEnding</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding)
		{
			 Factory.ExecuteMethod(this, "SaveAs2", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions, lineEnding });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.SaveAs2"/> </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object allowSubstitutions</param>
		/// <param name="lineEnding">optional object lineEnding</param>
		/// <param name="addBiDiMarks">optional object addBiDiMarks</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding, object addBiDiMarks)
		{
			 Factory.ExecuteMethod(this, "SaveAs2", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions, lineEnding, addBiDiMarks });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.SetCompatibilityMode"/> </remarks>
		/// <param name="mode">Int32 mode</param>
		[SupportByVersion("Word", 14,15,16)]
		public void SetCompatibilityMode(Int32 mode)
		{
			 Factory.ExecuteMethod(this, "SetCompatibilityMode", mode);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.document.returntolastreadposition"/> </remarks>
		[SupportByVersion("Word", 15, 16)]
		public Int32 ReturnToLastReadPosition()
		{
			return Factory.ExecuteInt32MethodGet(this, "ReturnToLastReadPosition");
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object allowSubstitutions</param>
		/// <param name="lineEnding">optional object lineEnding</param>
		/// <param name="addBiDiMarks">optional object addBiDiMarks</param>
		/// <param name="compatibilityMode">optional object compatibilityMode</param>
		[SupportByVersion("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding, object addBiDiMarks, object compatibilityMode)
		{
			 Factory.ExecuteMethod(this, "SaveCopyAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions, lineEnding, addBiDiMarks, compatibilityMode });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public void SaveCopyAs()
		{
			 Factory.ExecuteMethod(this, "SaveCopyAs");
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public void SaveCopyAs(object fileName)
		{
			 Factory.ExecuteMethod(this, "SaveCopyAs", fileName);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat)
		{
			 Factory.ExecuteMethod(this, "SaveCopyAs", fileName, fileFormat);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments)
		{
			 Factory.ExecuteMethod(this, "SaveCopyAs", fileName, fileFormat, lockComments);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password)
		{
			 Factory.ExecuteMethod(this, "SaveCopyAs", fileName, fileFormat, lockComments, password);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles)
		{
			 Factory.ExecuteMethod(this, "SaveCopyAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword)
		{
			 Factory.ExecuteMethod(this, "SaveCopyAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended)
		{
			 Factory.ExecuteMethod(this, "SaveCopyAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts)
		{
			 Factory.ExecuteMethod(this, "SaveCopyAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat)
		{
			 Factory.ExecuteMethod(this, "SaveCopyAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData)
		{
			 Factory.ExecuteMethod(this, "SaveCopyAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter)
		{
			 Factory.ExecuteMethod(this, "SaveCopyAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding)
		{
			 Factory.ExecuteMethod(this, "SaveCopyAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks)
		{
			 Factory.ExecuteMethod(this, "SaveCopyAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object allowSubstitutions</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions)
		{
			 Factory.ExecuteMethod(this, "SaveCopyAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object allowSubstitutions</param>
		/// <param name="lineEnding">optional object lineEnding</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding)
		{
			 Factory.ExecuteMethod(this, "SaveCopyAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions, lineEnding });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="embedTrueTypeFonts">optional object embedTrueTypeFonts</param>
		/// <param name="saveNativePictureFormat">optional object saveNativePictureFormat</param>
		/// <param name="saveFormsData">optional object saveFormsData</param>
		/// <param name="saveAsAOCELetter">optional object saveAsAOCELetter</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="insertLineBreaks">optional object insertLineBreaks</param>
		/// <param name="allowSubstitutions">optional object allowSubstitutions</param>
		/// <param name="lineEnding">optional object lineEnding</param>
		/// <param name="addBiDiMarks">optional object addBiDiMarks</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding, object addBiDiMarks)
		{
			 Factory.ExecuteMethod(this, "SaveCopyAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions, lineEnding, addBiDiMarks });
		}

		#endregion

		#pragma warning restore
	}
}
