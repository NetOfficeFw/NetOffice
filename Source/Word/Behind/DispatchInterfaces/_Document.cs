using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.WordApi;

namespace NetOffice.WordApi.Behind
{
	/// <summary>
	/// DispatchInterface _Document 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Document : COMObject, NetOffice.WordApi._Document
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
                    _contractType = typeof(NetOffice.WordApi._Document);
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196900.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual string Name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822944.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", typeof(NetOffice.WordApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845529.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Creator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840549.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196862.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object BuiltInDocumentProperties
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "BuiltInDocumentProperties");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195603.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object CustomDocumentProperties
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "CustomDocumentProperties");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821867.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual string Path
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Path");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194977.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Bookmarks Bookmarks
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Bookmarks>(this, "Bookmarks", typeof(NetOffice.WordApi.Bookmarks));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835455.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Tables Tables
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Tables>(this, "Tables", typeof(NetOffice.WordApi.Tables));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197126.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Footnotes Footnotes
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Footnotes>(this, "Footnotes", typeof(NetOffice.WordApi.Footnotes));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194032.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Endnotes Endnotes
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Endnotes>(this, "Endnotes", typeof(NetOffice.WordApi.Endnotes));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845880.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Comments Comments
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Comments>(this, "Comments", typeof(NetOffice.WordApi.Comments));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823228.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdDocumentType Type
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdDocumentType>(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191749.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool AutoHyphenation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AutoHyphenation");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AutoHyphenation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845783.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool HyphenateCaps
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HyphenateCaps");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HyphenateCaps", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193110.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 HyphenationZone
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "HyphenationZone");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HyphenationZone", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820862.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 ConsecutiveHyphensLimit
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ConsecutiveHyphensLimit");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ConsecutiveHyphensLimit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822125.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Sections Sections
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Sections>(this, "Sections", typeof(NetOffice.WordApi.Sections));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836325.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Paragraphs Paragraphs
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Paragraphs>(this, "Paragraphs", typeof(NetOffice.WordApi.Paragraphs));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845024.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Words Words
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Words>(this, "Words", typeof(NetOffice.WordApi.Words));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194403.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Sentences Sentences
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Sentences>(this, "Sentences", typeof(NetOffice.WordApi.Sentences));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191729.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Characters Characters
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Characters>(this, "Characters", typeof(NetOffice.WordApi.Characters));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821229.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Fields Fields
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Fields>(this, "Fields", typeof(NetOffice.WordApi.Fields));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840117.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.FormFields FormFields
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.FormFields>(this, "FormFields", typeof(NetOffice.WordApi.FormFields));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193100.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Styles Styles
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Styles>(this, "Styles", typeof(NetOffice.WordApi.Styles));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197117.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Frames Frames
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Frames>(this, "Frames", typeof(NetOffice.WordApi.Frames));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191950.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TablesOfFigures TablesOfFigures
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.TablesOfFigures>(this, "TablesOfFigures", typeof(NetOffice.WordApi.TablesOfFigures));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834524.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Variables Variables
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Variables>(this, "Variables", typeof(NetOffice.WordApi.Variables));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198370.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMerge MailMerge
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.MailMerge>(this, "MailMerge", typeof(NetOffice.WordApi.MailMerge));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844798.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Envelope Envelope
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Envelope>(this, "Envelope", typeof(NetOffice.WordApi.Envelope));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821285.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual string FullName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FullName");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192540.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Revisions Revisions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Revisions>(this, "Revisions", typeof(NetOffice.WordApi.Revisions));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822932.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TablesOfContents TablesOfContents
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.TablesOfContents>(this, "TablesOfContents", typeof(NetOffice.WordApi.TablesOfContents));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837912.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TablesOfAuthorities TablesOfAuthorities
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.TablesOfAuthorities>(this, "TablesOfAuthorities", typeof(NetOffice.WordApi.TablesOfAuthorities));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839306.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.PageSetup PageSetup
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.PageSetup>(this, "PageSetup", typeof(NetOffice.WordApi.PageSetup));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "PageSetup", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837336.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Windows Windows
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Windows>(this, "Windows", typeof(NetOffice.WordApi.Windows));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool HasRoutingSlip
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasRoutingSlip");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HasRoutingSlip", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.RoutingSlip RoutingSlip
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.RoutingSlip>(this, "RoutingSlip", typeof(NetOffice.WordApi.RoutingSlip));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool Routed
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Routed");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838095.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TablesOfAuthoritiesCategories TablesOfAuthoritiesCategories
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.TablesOfAuthoritiesCategories>(this, "TablesOfAuthoritiesCategories", typeof(NetOffice.WordApi.TablesOfAuthoritiesCategories));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194976.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Indexes Indexes
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Indexes>(this, "Indexes", typeof(NetOffice.WordApi.Indexes));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194753.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool Saved
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Saved");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Saved", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821853.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Range Content
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "Content", typeof(NetOffice.WordApi.Range));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198228.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Window ActiveWindow
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Window>(this, "ActiveWindow", typeof(NetOffice.WordApi.Window));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192728.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdDocumentKind Kind
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdDocumentKind>(this, "Kind");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Kind", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196223.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool ReadOnly
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ReadOnly");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195362.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Subdocuments Subdocuments
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Subdocuments>(this, "Subdocuments", typeof(NetOffice.WordApi.Subdocuments));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840840.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool IsMasterDocument
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsMasterDocument");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196079.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single DefaultTabStop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "DefaultTabStop");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultTabStop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836281.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool EmbedTrueTypeFonts
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EmbedTrueTypeFonts");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EmbedTrueTypeFonts", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845567.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool SaveFormsData
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SaveFormsData");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SaveFormsData", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838914.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool ReadOnlyRecommended
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ReadOnlyRecommended");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ReadOnlyRecommended", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844828.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool SaveSubsetFonts
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SaveSubsetFonts");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SaveSubsetFonts", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840506.aspx </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdCompatibility type</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool get_Compatibility(NetOffice.WordApi.Enums.WdCompatibility type)
		{
			return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Compatibility", type);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdCompatibility type</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual void set_Compatibility(NetOffice.WordApi.Enums.WdCompatibility type, bool value)
		{
			InvokerService.InvokeInternal.ExecutePropertySet(this, "Compatibility", type, value);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Compatibility
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840506.aspx </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdCompatibility type</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), Redirect("get_Compatibility")]
		public virtual bool Compatibility(NetOffice.WordApi.Enums.WdCompatibility type)
		{
			return get_Compatibility(type);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197823.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.StoryRanges StoryRanges
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.StoryRanges>(this, "StoryRanges", typeof(NetOffice.WordApi.StoryRanges));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821872.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.CommandBars CommandBars
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBars>(this, "CommandBars", typeof(NetOffice.OfficeApi.CommandBars));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192771.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool IsSubdocument
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsSubdocument");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840755.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 SaveFormat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SaveFormat");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836643.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdProtectionType ProtectionType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdProtectionType>(this, "ProtectionType");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837239.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Hyperlinks Hyperlinks
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Hyperlinks>(this, "Hyperlinks", typeof(NetOffice.WordApi.Hyperlinks));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197211.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Shapes Shapes
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Shapes>(this, "Shapes", typeof(NetOffice.WordApi.Shapes));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839163.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.ListTemplates ListTemplates
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ListTemplates>(this, "ListTemplates", typeof(NetOffice.WordApi.ListTemplates));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845137.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Lists Lists
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Lists>(this, "Lists", typeof(NetOffice.WordApi.Lists));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821398.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool UpdateStylesOnOpen
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "UpdateStylesOnOpen");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UpdateStylesOnOpen", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839734.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object AttachedTemplate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "AttachedTemplate");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "AttachedTemplate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844996.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.InlineShapes InlineShapes
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.InlineShapes>(this, "InlineShapes", typeof(NetOffice.WordApi.InlineShapes));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844976.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Shape Background
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Shape>(this, "Background", typeof(NetOffice.WordApi.Shape));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Background", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193109.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool GrammarChecked
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "GrammarChecked");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GrammarChecked", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845040.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool SpellingChecked
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SpellingChecked");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SpellingChecked", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836692.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool ShowGrammaticalErrors
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowGrammaticalErrors");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowGrammaticalErrors", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821056.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool ShowSpellingErrors
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowSpellingErrors");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowSpellingErrors", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Versions Versions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Versions>(this, "Versions", typeof(NetOffice.WordApi.Versions));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool ShowSummary
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowSummary");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowSummary", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdSummaryMode SummaryViewMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdSummaryMode>(this, "SummaryViewMode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "SummaryViewMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 SummaryLength
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SummaryLength");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SummaryLength", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool PrintFractionalWidths
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PrintFractionalWidths");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PrintFractionalWidths", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196987.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool PrintPostScriptOverText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PrintPostScriptOverText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PrintPostScriptOverText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840423.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object Container
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Container");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838735.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool PrintFormsData
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PrintFormsData");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PrintFormsData", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198090.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.ListParagraphs ListParagraphs
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ListParagraphs>(this, "ListParagraphs", typeof(NetOffice.WordApi.ListParagraphs));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192387.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual string Password
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Password");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Password", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839518.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual string WritePassword
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "WritePassword");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "WritePassword", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194500.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool HasPassword
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasPassword");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837527.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool WriteReserved
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "WriteReserved");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844946.aspx </remarks>
		/// <param name="languageID">object languageID</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string get_ActiveWritingStyle(object languageID)
		{
			return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ActiveWritingStyle", languageID);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="languageID">object languageID</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual void set_ActiveWritingStyle(object languageID, string value)
		{
			InvokerService.InvokeInternal.ExecutePropertySet(this, "ActiveWritingStyle", languageID, value);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_ActiveWritingStyle
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844946.aspx </remarks>
		/// <param name="languageID">object languageID</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), Redirect("get_ActiveWritingStyle")]
		public virtual string ActiveWritingStyle(object languageID)
		{
			return get_ActiveWritingStyle(languageID);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193401.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool UserControl
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "UserControl");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UserControl", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool HasMailer
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasMailer");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HasMailer", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Mailer Mailer
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Mailer>(this, "Mailer", typeof(NetOffice.WordApi.Mailer));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839868.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.ReadabilityStatistics ReadabilityStatistics
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ReadabilityStatistics>(this, "ReadabilityStatistics", typeof(NetOffice.WordApi.ReadabilityStatistics));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192400.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.ProofreadingErrors GrammaticalErrors
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ProofreadingErrors>(this, "GrammaticalErrors", typeof(NetOffice.WordApi.ProofreadingErrors));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838118.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.ProofreadingErrors SpellingErrors
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ProofreadingErrors>(this, "SpellingErrors", typeof(NetOffice.WordApi.ProofreadingErrors));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837668.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.VBIDEApi.VBProject VBProject
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBProject>(this, "VBProject", typeof(NetOffice.VBIDEApi.VBProject));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840586.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool FormsDesign
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "FormsDesign");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string _CodeName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "_CodeName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "_CodeName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197577.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual string CodeName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CodeName");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821373.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool SnapToGrid
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SnapToGrid");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SnapToGrid", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837193.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool SnapToShapes
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SnapToShapes");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SnapToShapes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839124.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single GridDistanceHorizontal
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "GridDistanceHorizontal");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridDistanceHorizontal", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195287.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single GridDistanceVertical
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "GridDistanceVertical");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridDistanceVertical", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839558.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single GridOriginHorizontal
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "GridOriginHorizontal");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridOriginHorizontal", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198193.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single GridOriginVertical
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "GridOriginVertical");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridOriginVertical", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821306.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 GridSpaceBetweenHorizontalLines
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "GridSpaceBetweenHorizontalLines");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridSpaceBetweenHorizontalLines", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821996.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 GridSpaceBetweenVerticalLines
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "GridSpaceBetweenVerticalLines");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridSpaceBetweenVerticalLines", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845752.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool GridOriginFromMargin
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "GridOriginFromMargin");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridOriginFromMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836931.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool KerningByAlgorithm
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "KerningByAlgorithm");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "KerningByAlgorithm", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191748.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdJustificationMode JustificationMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdJustificationMode>(this, "JustificationMode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "JustificationMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845667.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdFarEastLineBreakLevel FarEastLineBreakLevel
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdFarEastLineBreakLevel>(this, "FarEastLineBreakLevel");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "FarEastLineBreakLevel", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844966.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual string NoLineBreakBefore
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "NoLineBreakBefore");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "NoLineBreakBefore", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192597.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual string NoLineBreakAfter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "NoLineBreakAfter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "NoLineBreakAfter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838067.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool TrackRevisions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "TrackRevisions");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TrackRevisions", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192825.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool PrintRevisions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PrintRevisions");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PrintRevisions", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool ShowRevisions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowRevisions");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowRevisions", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192741.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual string ActiveTheme
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ActiveTheme");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837037.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual string ActiveThemeDisplayName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ActiveThemeDisplayName");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839292.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Email Email
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Email>(this, "Email", typeof(NetOffice.WordApi.Email));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196093.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.Scripts Scripts
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Scripts>(this, "Scripts", typeof(NetOffice.OfficeApi.Scripts));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191794.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool LanguageDetected
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "LanguageDetected");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LanguageDetected", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838486.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdFarEastLineBreakLanguageID FarEastLineBreakLanguage
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdFarEastLineBreakLanguageID>(this, "FarEastLineBreakLanguage");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "FarEastLineBreakLanguage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194305.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Frameset Frameset
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Frameset>(this, "Frameset", typeof(NetOffice.WordApi.Frameset));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839615.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object ClickAndTypeParagraphStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ClickAndTypeParagraphStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ClickAndTypeParagraphStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.HTMLProject HTMLProject
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.HTMLProject>(this, "HTMLProject", typeof(NetOffice.OfficeApi.HTMLProject));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844954.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.WebOptions WebOptions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.WebOptions>(this, "WebOptions", typeof(NetOffice.WordApi.WebOptions));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835467.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoEncoding OpenEncoding
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoEncoding>(this, "OpenEncoding");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834893.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoEncoding SaveEncoding
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoEncoding>(this, "SaveEncoding");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "SaveEncoding", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835162.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool OptimizeForWord97
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "OptimizeForWord97");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OptimizeForWord97", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836069.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool VBASigned
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "VBASigned");
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840465.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.MsoEnvelope MailEnvelope
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.MsoEnvelope>(this, "MailEnvelope", typeof(NetOffice.OfficeApi.MsoEnvelope));
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194348.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual bool DisableFeatures
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisableFeatures");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisableFeatures", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194604.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual bool DoNotEmbedSystemFonts
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DoNotEmbedSystemFonts");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DoNotEmbedSystemFonts", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193069.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.SignatureSet Signatures
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SignatureSet>(this, "Signatures", typeof(NetOffice.OfficeApi.SignatureSet));
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194661.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual string DefaultTargetFrame
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DefaultTargetFrame");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultTargetFrame", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822985.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.HTMLDivisions HTMLDivisions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.HTMLDivisions>(this, "HTMLDivisions", typeof(NetOffice.WordApi.HTMLDivisions));
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196211.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdDisableFeaturesIntroducedAfter DisableFeaturesIntroducedAfter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdDisableFeaturesIntroducedAfter>(this, "DisableFeaturesIntroducedAfter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "DisableFeaturesIntroducedAfter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838361.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
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
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.SmartTags SmartTags
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.SmartTags>(this, "SmartTags", typeof(NetOffice.WordApi.SmartTags));
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual bool EmbedSmartTags
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EmbedSmartTags");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EmbedSmartTags", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual bool SmartTagsAsXMLProps
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SmartTagsAsXMLProps");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SmartTagsAsXMLProps", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835460.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoEncoding TextEncoding
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoEncoding>(this, "TextEncoding");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "TextEncoding", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198078.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdLineEndingType TextLineEnding
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdLineEndingType>(this, "TextLineEnding");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "TextLineEnding", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845673.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.StyleSheets StyleSheets
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.StyleSheets>(this, "StyleSheets", typeof(NetOffice.WordApi.StyleSheets));
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837042.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual object DefaultTableStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "DefaultTableStyle");
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194870.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual string PasswordEncryptionProvider
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "PasswordEncryptionProvider");
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195788.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual string PasswordEncryptionAlgorithm
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "PasswordEncryptionAlgorithm");
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193119.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual Int32 PasswordEncryptionKeyLength
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "PasswordEncryptionKeyLength");
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822966.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual bool PasswordEncryptionFileProperties
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PasswordEncryptionFileProperties");
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836336.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual bool EmbedLinguisticData
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EmbedLinguisticData");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EmbedLinguisticData", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839893.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual bool FormattingShowFont
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "FormattingShowFont");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FormattingShowFont", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839706.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual bool FormattingShowClear
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "FormattingShowClear");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FormattingShowClear", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836749.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual bool FormattingShowParagraph
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "FormattingShowParagraph");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FormattingShowParagraph", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193041.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual bool FormattingShowNumbering
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "FormattingShowNumbering");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FormattingShowNumbering", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194361.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdShowFilter FormattingShowFilter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdShowFilter>(this, "FormattingShowFilter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "FormattingShowFilter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191744.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.Permission Permission
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Permission>(this, "Permission", typeof(NetOffice.OfficeApi.Permission));
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual NetOffice.WordApi.XMLNodes XMLNodes
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLNodes>(this, "XMLNodes", typeof(NetOffice.WordApi.XMLNodes));
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198201.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual NetOffice.WordApi.XMLSchemaReferences XMLSchemaReferences
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLSchemaReferences>(this, "XMLSchemaReferences", typeof(NetOffice.WordApi.XMLSchemaReferences));
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840776.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.SmartDocument SmartDocument
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartDocument>(this, "SmartDocument", typeof(NetOffice.OfficeApi.SmartDocument));
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.SharedWorkspace SharedWorkspace
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SharedWorkspace>(this, "SharedWorkspace", typeof(NetOffice.OfficeApi.SharedWorkspace));
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837910.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.Sync Sync
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Sync>(this, "Sync", typeof(NetOffice.OfficeApi.Sync));
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838344.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual bool EnforceStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EnforceStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnforceStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822185.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual bool AutoFormatOverride
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AutoFormatOverride");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AutoFormatOverride", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual bool XMLSaveDataOnly
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "XMLSaveDataOnly");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "XMLSaveDataOnly", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual bool XMLHideNamespaces
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "XMLHideNamespaces");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "XMLHideNamespaces", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196205.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual bool XMLShowAdvancedErrors
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "XMLShowAdvancedErrors");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "XMLShowAdvancedErrors", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836689.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual bool XMLUseXSLTWhenSaving
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "XMLUseXSLTWhenSaving");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "XMLUseXSLTWhenSaving", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838300.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual string XMLSaveThroughXSLT
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "XMLSaveThroughXSLT");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "XMLSaveThroughXSLT", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191946.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.DocumentLibraryVersions DocumentLibraryVersions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.DocumentLibraryVersions>(this, "DocumentLibraryVersions", typeof(NetOffice.OfficeApi.DocumentLibraryVersions));
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196654.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual bool ReadingModeLayoutFrozen
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ReadingModeLayoutFrozen");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ReadingModeLayoutFrozen", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194610.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual bool RemoveDateAndTime
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "RemoveDateAndTime");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RemoveDateAndTime", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual NetOffice.WordApi.XMLChildNodeSuggestions ChildNodeSuggestions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLChildNodeSuggestions>(this, "ChildNodeSuggestions", typeof(NetOffice.WordApi.XMLChildNodeSuggestions));
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual NetOffice.WordApi.XMLNodes XMLSchemaViolations
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLNodes>(this, "XMLSchemaViolations", typeof(NetOffice.WordApi.XMLNodes));
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191938.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual Int32 ReadingLayoutSizeX
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ReadingLayoutSizeX");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ReadingLayoutSizeX", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839167.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual Int32 ReadingLayoutSizeY
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ReadingLayoutSizeY");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ReadingLayoutSizeY", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191767.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdStyleSort StyleSortMethod
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdStyleSort>(this, "StyleSortMethod");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "StyleSortMethod", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844919.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.OfficeApi.MetaProperties ContentTypeProperties
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.MetaProperties>(this, "ContentTypeProperties", typeof(NetOffice.OfficeApi.MetaProperties));
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197907.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual bool TrackMoves
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "TrackMoves");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TrackMoves", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836881.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual bool TrackFormatting
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "TrackFormatting");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TrackFormatting", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Dummy1
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Dummy1");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837488.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.OMaths OMaths
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMaths>(this, "OMaths", typeof(NetOffice.WordApi.OMaths));
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Dummy3
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Dummy3");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839289.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.OfficeApi.ServerPolicy ServerPolicy
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ServerPolicy>(this, "ServerPolicy", typeof(NetOffice.OfficeApi.ServerPolicy));
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822382.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.ContentControls ContentControls
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ContentControls>(this, "ContentControls", typeof(NetOffice.WordApi.ContentControls));
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839144.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.OfficeApi.DocumentInspectors DocumentInspectors
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.DocumentInspectors>(this, "DocumentInspectors", typeof(NetOffice.OfficeApi.DocumentInspectors));
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834552.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.Bibliography Bibliography
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Bibliography>(this, "Bibliography", typeof(NetOffice.WordApi.Bibliography));
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198209.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual bool LockTheme
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "LockTheme");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LockTheme", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839340.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual bool LockQuickStyleSet
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "LockQuickStyleSet");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LockQuickStyleSet", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821063.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual string OriginalDocumentTitle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OriginalDocumentTitle");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834817.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual string RevisedDocumentTitle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "RevisedDocumentTitle");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193091.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.OfficeApi.CustomXMLParts CustomXMLParts
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLParts>(this, "CustomXMLParts", typeof(NetOffice.OfficeApi.CustomXMLParts));
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195284.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual bool FormattingShowNextLevel
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "FormattingShowNextLevel");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FormattingShowNextLevel", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191723.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual bool FormattingShowUserStyleName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "FormattingShowUserStyleName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FormattingShowUserStyleName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822952.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.Research Research
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Research>(this, "Research", typeof(NetOffice.WordApi.Research));
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838930.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual bool Final
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Final");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Final", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821662.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdOMathBreakBin OMathBreakBin
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdOMathBreakBin>(this, "OMathBreakBin");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "OMathBreakBin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835681.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdOMathBreakSub OMathBreakSub
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdOMathBreakSub>(this, "OMathBreakSub");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "OMathBreakSub", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196528.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdOMathJc OMathJc
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdOMathJc>(this, "OMathJc");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "OMathJc", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195080.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual Single OMathLeftMargin
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "OMathLeftMargin");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OMathLeftMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192826.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual Single OMathRightMargin
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "OMathRightMargin");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OMathRightMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195018.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual Single OMathWrap
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "OMathWrap");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OMathWrap", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822912.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual bool OMathIntSubSupLim
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "OMathIntSubSupLim");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OMathIntSubSupLim", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192808.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual bool OMathNarySupSubLim
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "OMathNarySupSubLim");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OMathNarySupSubLim", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835679.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual bool OMathSmallFrac
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "OMathSmallFrac");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OMathSmallFrac", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197690.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual string WordOpenXML
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "WordOpenXML");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840566.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.OfficeApi.OfficeTheme DocumentTheme
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.OfficeTheme>(this, "DocumentTheme", typeof(NetOffice.OfficeApi.OfficeTheme));
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845747.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual bool HasVBProject
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasVBProject");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193851.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual string OMathFontName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OMathFontName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OMathFontName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836379.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual string EncryptionProvider
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EncryptionProvider");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EncryptionProvider", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838359.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual bool UseMathDefaults
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "UseMathDefaults");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UseMathDefaults", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195620.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual Int32 CurrentRsid
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CurrentRsid");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 DocID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "DocID");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196837.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public virtual Int32 CompatibilityMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CompatibilityMode");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837045.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public virtual NetOffice.WordApi.CoAuthoring CoAuthoring
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.CoAuthoring>(this, "CoAuthoring", typeof(NetOffice.WordApi.CoAuthoring));
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231858.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public virtual NetOffice.WordApi.Broadcast Broadcast
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Broadcast>(this, "Broadcast", typeof(NetOffice.WordApi.Broadcast));
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228844.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public virtual bool ChartDataPointTrack
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ChartDataPointTrack");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ChartDataPointTrack", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230857.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public virtual bool IsInAutosave
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsInAutosave");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196343.aspx </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		/// <param name="originalFormat">optional object originalFormat</param>
		/// <param name="routeDocument">optional object routeDocument</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Close(object saveChanges, object originalFormat, object routeDocument)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Close", saveChanges, originalFormat, routeDocument);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196343.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Close()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196343.aspx </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Close(object saveChanges)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Close", saveChanges);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196343.aspx </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		/// <param name="originalFormat">optional object originalFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Close(object saveChanges, object originalFormat)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Close", saveChanges, originalFormat);
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
		public virtual void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter });
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
		public virtual void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding, object addBiDiMarks)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions, lineEnding, addBiDiMarks });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void SaveAs()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void SaveAs(object fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", fileName);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void SaveAs(object fileName, object fileFormat)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", fileName, fileFormat);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void SaveAs(object fileName, object fileFormat, object lockComments)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", fileName, fileFormat, lockComments);
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
		public virtual void SaveAs(object fileName, object fileFormat, object lockComments, object password)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", fileName, fileFormat, lockComments, password);
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
		public virtual void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles });
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
		public virtual void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword });
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
		public virtual void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended });
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
		public virtual void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts });
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
		public virtual void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat });
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
		public virtual void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData });
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
		public virtual void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding });
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
		public virtual void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks });
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
		public virtual void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions });
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
		public virtual void SaveAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions, lineEnding });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821326.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Repaginate()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Repaginate");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822617.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void FitToPages()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FitToPages");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841098.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void ManualHyphenation()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ManualHyphenation");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845112.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Select()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Select");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845755.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void DataForm()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DataForm");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Route()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Route");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821625.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Save()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Save");
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
		public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOutOld()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOutOld(object background)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", background);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOutOld(object background, object append)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", background, append);
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
		public virtual void PrintOutOld(object background, object append, object range)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", background, append, range);
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
		public virtual void PrintOutOld(object background, object append, object range, object outputFileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", background, append, range, outputFileName);
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
		public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from });
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
		public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to });
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
		public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item });
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
		public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies });
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
		public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages });
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
		public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType });
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
		public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile });
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
		public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate });
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
		public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821630.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void SendMail()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendMail");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821608.aspx </remarks>
		/// <param name="start">optional object start</param>
		/// <param name="end">optional object end</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Range Range(object start, object end)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "Range", typeof(NetOffice.WordApi.Range), start, end);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821608.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Range Range()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "Range", typeof(NetOffice.WordApi.Range));
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821608.aspx </remarks>
		/// <param name="start">optional object start</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Range Range(object start)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "Range", typeof(NetOffice.WordApi.Range), start);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823210.aspx </remarks>
		/// <param name="which">NetOffice.WordApi.Enums.WdAutoMacros which</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void RunAutoMacro(NetOffice.WordApi.Enums.WdAutoMacros which)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RunAutoMacro", which);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822131.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Activate()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Activate");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195898.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintPreview()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintPreview");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836585.aspx </remarks>
		/// <param name="what">optional object what</param>
		/// <param name="which">optional object which</param>
		/// <param name="count">optional object count</param>
		/// <param name="name">optional object name</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Range GoTo(object what, object which, object count, object name)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoTo", typeof(NetOffice.WordApi.Range), what, which, count, name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836585.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Range GoTo()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoTo", typeof(NetOffice.WordApi.Range));
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836585.aspx </remarks>
		/// <param name="what">optional object what</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Range GoTo(object what)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoTo", typeof(NetOffice.WordApi.Range), what);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836585.aspx </remarks>
		/// <param name="what">optional object what</param>
		/// <param name="which">optional object which</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Range GoTo(object what, object which)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoTo", typeof(NetOffice.WordApi.Range), what, which);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836585.aspx </remarks>
		/// <param name="what">optional object what</param>
		/// <param name="which">optional object which</param>
		/// <param name="count">optional object count</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Range GoTo(object what, object which, object count)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoTo", typeof(NetOffice.WordApi.Range), what, which, count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840796.aspx </remarks>
		/// <param name="times">optional object times</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool Undo(object times)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Undo", times);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840796.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool Undo()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Undo");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845577.aspx </remarks>
		/// <param name="times">optional object times</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool Redo(object times)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Redo", times);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845577.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool Redo()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Redo");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840638.aspx </remarks>
		/// <param name="statistic">NetOffice.WordApi.Enums.WdStatistic statistic</param>
		/// <param name="includeFootnotesAndEndnotes">optional object includeFootnotesAndEndnotes</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 ComputeStatistics(NetOffice.WordApi.Enums.WdStatistic statistic, object includeFootnotesAndEndnotes)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ComputeStatistics", statistic, includeFootnotesAndEndnotes);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840638.aspx </remarks>
		/// <param name="statistic">NetOffice.WordApi.Enums.WdStatistic statistic</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 ComputeStatistics(NetOffice.WordApi.Enums.WdStatistic statistic)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ComputeStatistics", statistic);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845133.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void MakeCompatibilityDefault()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MakeCompatibilityDefault");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230379.aspx </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType type</param>
		/// <param name="noReset">optional object noReset</param>
		/// <param name="password">optional object password</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Protect(NetOffice.WordApi.Enums.WdProtectionType type, object noReset, object password)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Protect", type, noReset, password);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230379.aspx </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType type</param>
		/// <param name="noReset">optional object noReset</param>
		/// <param name="password">optional object password</param>
		/// <param name="useIRM">optional object useIRM</param>
		/// <param name="enforceStyleLock">optional object enforceStyleLock</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void Protect(NetOffice.WordApi.Enums.WdProtectionType type, object noReset, object password, object useIRM, object enforceStyleLock)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Protect", new object[]{ type, noReset, password, useIRM, enforceStyleLock });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230379.aspx </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType type</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Protect(NetOffice.WordApi.Enums.WdProtectionType type)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Protect", type);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230379.aspx </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType type</param>
		/// <param name="noReset">optional object noReset</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Protect(NetOffice.WordApi.Enums.WdProtectionType type, object noReset)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Protect", type, noReset);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230379.aspx </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType type</param>
		/// <param name="noReset">optional object noReset</param>
		/// <param name="password">optional object password</param>
		/// <param name="useIRM">optional object useIRM</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void Protect(NetOffice.WordApi.Enums.WdProtectionType type, object noReset, object password, object useIRM)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Protect", type, noReset, password, useIRM);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845016.aspx </remarks>
		/// <param name="password">optional object password</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Unprotect(object password)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Unprotect", password);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845016.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Unprotect()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Unprotect");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdEditionType type</param>
		/// <param name="option">NetOffice.WordApi.Enums.WdEditionOption option</param>
		/// <param name="name">string name</param>
		/// <param name="format">optional object format</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void EditionOptions(NetOffice.WordApi.Enums.WdEditionType type, NetOffice.WordApi.Enums.WdEditionOption option, string name, object format)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EditionOptions", type, option, name, format);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdEditionType type</param>
		/// <param name="option">NetOffice.WordApi.Enums.WdEditionOption option</param>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void EditionOptions(NetOffice.WordApi.Enums.WdEditionType type, NetOffice.WordApi.Enums.WdEditionOption option, string name)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EditionOptions", type, option, name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821600.aspx </remarks>
		/// <param name="letterContent">optional object letterContent</param>
		/// <param name="wizardMode">optional object wizardMode</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void RunLetterWizard(object letterContent, object wizardMode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RunLetterWizard", letterContent, wizardMode);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821600.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void RunLetterWizard()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RunLetterWizard");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821600.aspx </remarks>
		/// <param name="letterContent">optional object letterContent</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void RunLetterWizard(object letterContent)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RunLetterWizard", letterContent);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836106.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.LetterContent GetLetterContent()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.LetterContent>(this, "GetLetterContent", typeof(NetOffice.WordApi.LetterContent));
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822930.aspx </remarks>
		/// <param name="letterContent">object letterContent</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void SetLetterContent(object letterContent)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetLetterContent", letterContent);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840260.aspx </remarks>
		/// <param name="template">string template</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void CopyStylesFromTemplate(string template)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CopyStylesFromTemplate", template);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840983.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void UpdateStyles()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateStyles");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834835.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void CheckGrammar()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckGrammar");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx </remarks>
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
		public virtual void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9, object customDictionary10)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9, customDictionary10 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void CheckSpelling()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckSpelling");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void CheckSpelling(object customDictionary)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckSpelling", customDictionary);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void CheckSpelling(object customDictionary, object ignoreUppercase)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckSpelling", customDictionary, ignoreUppercase);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckSpelling", customDictionary, ignoreUppercase, alwaysSuggest);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckSpelling", customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx </remarks>
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
		public virtual void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx </remarks>
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
		public virtual void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx </remarks>
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
		public virtual void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835796.aspx </remarks>
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
		public virtual void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840237.aspx </remarks>
		/// <param name="address">optional object address</param>
		/// <param name="subAddress">optional object subAddress</param>
		/// <param name="newWindow">optional object newWindow</param>
		/// <param name="addHistory">optional object addHistory</param>
		/// <param name="extraInfo">optional object extraInfo</param>
		/// <param name="method">optional object method</param>
		/// <param name="headerInfo">optional object headerInfo</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void FollowHyperlink(object address, object subAddress, object newWindow, object addHistory, object extraInfo, object method, object headerInfo)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", new object[]{ address, subAddress, newWindow, addHistory, extraInfo, method, headerInfo });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840237.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void FollowHyperlink()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840237.aspx </remarks>
		/// <param name="address">optional object address</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void FollowHyperlink(object address)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", address);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840237.aspx </remarks>
		/// <param name="address">optional object address</param>
		/// <param name="subAddress">optional object subAddress</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void FollowHyperlink(object address, object subAddress)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", address, subAddress);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840237.aspx </remarks>
		/// <param name="address">optional object address</param>
		/// <param name="subAddress">optional object subAddress</param>
		/// <param name="newWindow">optional object newWindow</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void FollowHyperlink(object address, object subAddress, object newWindow)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", address, subAddress, newWindow);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840237.aspx </remarks>
		/// <param name="address">optional object address</param>
		/// <param name="subAddress">optional object subAddress</param>
		/// <param name="newWindow">optional object newWindow</param>
		/// <param name="addHistory">optional object addHistory</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void FollowHyperlink(object address, object subAddress, object newWindow, object addHistory)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", address, subAddress, newWindow, addHistory);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840237.aspx </remarks>
		/// <param name="address">optional object address</param>
		/// <param name="subAddress">optional object subAddress</param>
		/// <param name="newWindow">optional object newWindow</param>
		/// <param name="addHistory">optional object addHistory</param>
		/// <param name="extraInfo">optional object extraInfo</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void FollowHyperlink(object address, object subAddress, object newWindow, object addHistory, object extraInfo)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", new object[]{ address, subAddress, newWindow, addHistory, extraInfo });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840237.aspx </remarks>
		/// <param name="address">optional object address</param>
		/// <param name="subAddress">optional object subAddress</param>
		/// <param name="newWindow">optional object newWindow</param>
		/// <param name="addHistory">optional object addHistory</param>
		/// <param name="extraInfo">optional object extraInfo</param>
		/// <param name="method">optional object method</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void FollowHyperlink(object address, object subAddress, object newWindow, object addHistory, object extraInfo, object method)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", new object[]{ address, subAddress, newWindow, addHistory, extraInfo, method });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839781.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void AddToFavorites()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddToFavorites");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195614.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Reload()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Reload");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="length">optional object length</param>
		/// <param name="mode">optional object mode</param>
		/// <param name="updateProperties">optional object updateProperties</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Range AutoSummarize(object length, object mode, object updateProperties)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "AutoSummarize", typeof(NetOffice.WordApi.Range), length, mode, updateProperties);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Range AutoSummarize()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "AutoSummarize", typeof(NetOffice.WordApi.Range));
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="length">optional object length</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Range AutoSummarize(object length)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "AutoSummarize", typeof(NetOffice.WordApi.Range), length);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="length">optional object length</param>
		/// <param name="mode">optional object mode</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Range AutoSummarize(object length, object mode)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "AutoSummarize", typeof(NetOffice.WordApi.Range), length, mode);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193060.aspx </remarks>
		/// <param name="numberType">optional object numberType</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void RemoveNumbers(object numberType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RemoveNumbers", numberType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193060.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void RemoveNumbers()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RemoveNumbers");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838874.aspx </remarks>
		/// <param name="numberType">optional object numberType</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void ConvertNumbersToText(object numberType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ConvertNumbersToText", numberType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838874.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void ConvertNumbersToText()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ConvertNumbersToText");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836858.aspx </remarks>
		/// <param name="numberType">optional object numberType</param>
		/// <param name="level">optional object level</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 CountNumberedItems(object numberType, object level)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "CountNumberedItems", numberType, level);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836858.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 CountNumberedItems()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "CountNumberedItems");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836858.aspx </remarks>
		/// <param name="numberType">optional object numberType</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 CountNumberedItems(object numberType)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "CountNumberedItems", numberType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192151.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Post()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Post");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195394.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void ToggleFormsDesign()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ToggleFormsDesign");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192559.aspx </remarks>
		/// <param name="name">string name</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Compare(string name)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Compare", name);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192559.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		/// <param name="compareTarget">optional object compareTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		/// <param name="ignoreAllComparisonWarnings">optional object ignoreAllComparisonWarnings</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void Compare(string name, object authorName, object compareTarget, object detectFormatChanges, object ignoreAllComparisonWarnings, object addToRecentFiles)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Compare", new object[]{ name, authorName, compareTarget, detectFormatChanges, ignoreAllComparisonWarnings, addToRecentFiles });
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192559.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		/// <param name="compareTarget">optional object compareTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		/// <param name="ignoreAllComparisonWarnings">optional object ignoreAllComparisonWarnings</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="removePersonalInformation">optional object removePersonalInformation</param>
		/// <param name="removeDateAndTime">optional object removeDateAndTime</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void Compare(string name, object authorName, object compareTarget, object detectFormatChanges, object ignoreAllComparisonWarnings, object addToRecentFiles, object removePersonalInformation, object removeDateAndTime)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Compare", new object[]{ name, authorName, compareTarget, detectFormatChanges, ignoreAllComparisonWarnings, addToRecentFiles, removePersonalInformation, removeDateAndTime });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192559.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void Compare(string name, object authorName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Compare", name, authorName);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192559.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		/// <param name="compareTarget">optional object compareTarget</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void Compare(string name, object authorName, object compareTarget)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Compare", name, authorName, compareTarget);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192559.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		/// <param name="compareTarget">optional object compareTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void Compare(string name, object authorName, object compareTarget, object detectFormatChanges)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Compare", name, authorName, compareTarget, detectFormatChanges);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192559.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		/// <param name="compareTarget">optional object compareTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		/// <param name="ignoreAllComparisonWarnings">optional object ignoreAllComparisonWarnings</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void Compare(string name, object authorName, object compareTarget, object detectFormatChanges, object ignoreAllComparisonWarnings)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Compare", new object[]{ name, authorName, compareTarget, detectFormatChanges, ignoreAllComparisonWarnings });
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192559.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		/// <param name="compareTarget">optional object compareTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		/// <param name="ignoreAllComparisonWarnings">optional object ignoreAllComparisonWarnings</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="removePersonalInformation">optional object removePersonalInformation</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void Compare(string name, object authorName, object compareTarget, object detectFormatChanges, object ignoreAllComparisonWarnings, object addToRecentFiles, object removePersonalInformation)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Compare", new object[]{ name, authorName, compareTarget, detectFormatChanges, ignoreAllComparisonWarnings, addToRecentFiles, removePersonalInformation });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void UpdateSummaryProperties()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateSummaryProperties");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193699.aspx </remarks>
		/// <param name="referenceType">object referenceType</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual object GetCrossReferenceItems(object referenceType)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "GetCrossReferenceItems", referenceType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193992.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void AutoFormat()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AutoFormat");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837880.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void ViewCode()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ViewCode");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834519.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void ViewPropertyBrowser()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ViewPropertyBrowser");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void ForwardMailer()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ForwardMailer");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Reply()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Reply");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void ReplyAll()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ReplyAll");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="priority">optional object priority</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void SendMailer(object fileFormat, object priority)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendMailer", fileFormat, priority);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void SendMailer()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendMailer");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileFormat">optional object fileFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void SendMailer(object fileFormat)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendMailer", fileFormat);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195616.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void UndoClear()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "UndoClear");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192417.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PresentIt()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PresentIt");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838927.aspx </remarks>
		/// <param name="address">string address</param>
		/// <param name="subject">optional object subject</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void SendFax(string address, object subject)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendFax", address, subject);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838927.aspx </remarks>
		/// <param name="address">string address</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void SendFax(string address)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendFax", address);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839752.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Merge(string fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Merge", fileName);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839752.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="mergeTarget">optional object mergeTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		/// <param name="useFormattingFrom">optional object useFormattingFrom</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void Merge(string fileName, object mergeTarget, object detectFormatChanges, object useFormattingFrom, object addToRecentFiles)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Merge", new object[]{ fileName, mergeTarget, detectFormatChanges, useFormattingFrom, addToRecentFiles });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839752.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="mergeTarget">optional object mergeTarget</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void Merge(string fileName, object mergeTarget)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Merge", fileName, mergeTarget);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839752.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="mergeTarget">optional object mergeTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void Merge(string fileName, object mergeTarget, object detectFormatChanges)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Merge", fileName, mergeTarget, detectFormatChanges);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839752.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="mergeTarget">optional object mergeTarget</param>
		/// <param name="detectFormatChanges">optional object detectFormatChanges</param>
		/// <param name="useFormattingFrom">optional object useFormattingFrom</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void Merge(string fileName, object mergeTarget, object detectFormatChanges, object useFormattingFrom)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Merge", fileName, mergeTarget, detectFormatChanges, useFormattingFrom);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822702.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void ClosePrintPreview()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ClosePrintPreview");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834920.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void CheckConsistency()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckConsistency");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193392.aspx </remarks>
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
		public virtual NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode, object recipientGender, object returnAddressShortForm, object senderCity, object senderCode, object senderGender, object senderReference)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.LetterContent>(this, "CreateLetterContent", typeof(NetOffice.WordApi.LetterContent), new object[]{ dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cCList, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber, infoBlock, recipientCode, recipientGender, returnAddressShortForm, senderCity, senderCode, senderGender, senderReference });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193392.aspx </remarks>
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
		public virtual NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.LetterContent>(this, "CreateLetterContent", typeof(NetOffice.WordApi.LetterContent), new object[]{ dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cCList, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193392.aspx </remarks>
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
		public virtual NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.LetterContent>(this, "CreateLetterContent", typeof(NetOffice.WordApi.LetterContent), new object[]{ dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cCList, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber, infoBlock });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193392.aspx </remarks>
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
		public virtual NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.LetterContent>(this, "CreateLetterContent", typeof(NetOffice.WordApi.LetterContent), new object[]{ dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cCList, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber, infoBlock, recipientCode });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193392.aspx </remarks>
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
		public virtual NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode, object recipientGender)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.LetterContent>(this, "CreateLetterContent", typeof(NetOffice.WordApi.LetterContent), new object[]{ dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cCList, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber, infoBlock, recipientCode, recipientGender });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193392.aspx </remarks>
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
		public virtual NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode, object recipientGender, object returnAddressShortForm)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.LetterContent>(this, "CreateLetterContent", typeof(NetOffice.WordApi.LetterContent), new object[]{ dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cCList, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber, infoBlock, recipientCode, recipientGender, returnAddressShortForm });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193392.aspx </remarks>
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
		public virtual NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode, object recipientGender, object returnAddressShortForm, object senderCity)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.LetterContent>(this, "CreateLetterContent", typeof(NetOffice.WordApi.LetterContent), new object[]{ dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cCList, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber, infoBlock, recipientCode, recipientGender, returnAddressShortForm, senderCity });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193392.aspx </remarks>
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
		public virtual NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode, object recipientGender, object returnAddressShortForm, object senderCity, object senderCode)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.LetterContent>(this, "CreateLetterContent", typeof(NetOffice.WordApi.LetterContent), new object[]{ dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cCList, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber, infoBlock, recipientCode, recipientGender, returnAddressShortForm, senderCity, senderCode });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193392.aspx </remarks>
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
		public virtual NetOffice.WordApi.LetterContent CreateLetterContent(string dateFormat, bool includeHeaderFooter, string pageDesign, NetOffice.WordApi.Enums.WdLetterStyle letterStyle, bool letterhead, NetOffice.WordApi.Enums.WdLetterheadLocation letterheadLocation, Single letterheadSize, string recipientName, string recipientAddress, string salutation, NetOffice.WordApi.Enums.WdSalutationType salutationType, string recipientReference, string mailingInstructions, string attentionLine, string subject, string cCList, string returnAddress, string senderName, string closing, string senderCompany, string senderJobTitle, string senderInitials, Int32 enclosureNumber, object infoBlock, object recipientCode, object recipientGender, object returnAddressShortForm, object senderCity, object senderCode, object senderGender)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.LetterContent>(this, "CreateLetterContent", typeof(NetOffice.WordApi.LetterContent), new object[]{ dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cCList, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber, infoBlock, recipientCode, recipientGender, returnAddressShortForm, senderCity, senderCode, senderGender });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193342.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void AcceptAllRevisions()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AcceptAllRevisions");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838536.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void RejectAllRevisions()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RejectAllRevisions");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197127.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void DetectLanguage()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DetectLanguage");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835740.aspx </remarks>
		/// <param name="name">string name</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void ApplyTheme(string name)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyTheme", name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839088.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void RemoveTheme()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RemoveTheme");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835177.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void WebPagePreview()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "WebPagePreview");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195768.aspx </remarks>
		/// <param name="encoding">NetOffice.OfficeApi.Enums.MsoEncoding encoding</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void ReloadAs(NetOffice.OfficeApi.Enums.MsoEncoding encoding)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ReloadAs", encoding);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
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
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth, object printZoomPaperHeight)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth, printZoomPaperHeight });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
		/// <param name="background">optional object background</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object background)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", background);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object background, object append)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", background, append);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object background, object append, object range)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", background, append, range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object background, object append, object range, object outputFileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", background, append, range, outputFileName);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
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
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
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
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
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
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
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
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
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
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
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
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
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
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
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
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
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
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837331.aspx </remarks>
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
		public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="s">string s</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void sblt(string s)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "sblt", s);
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
		public virtual void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2000", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void SaveAs2000()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2000");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void SaveAs2000(object fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2000", fileName);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void SaveAs2000(object fileName, object fileFormat)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2000", fileName, fileFormat);
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
		public virtual void SaveAs2000(object fileName, object fileFormat, object lockComments)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2000", fileName, fileFormat, lockComments);
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
		public virtual void SaveAs2000(object fileName, object fileFormat, object lockComments, object password)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2000", fileName, fileFormat, lockComments, password);
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
		public virtual void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2000", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles });
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
		public virtual void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2000", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword });
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
		public virtual void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2000", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended });
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
		public virtual void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2000", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts });
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
		public virtual void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2000", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat });
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
		public virtual void SaveAs2000(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2000", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void Compare2000(string name)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Compare2000", name);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void Merge2000(string fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Merge2000", fileName);
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
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth, object printZoomPaperHeight)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth, printZoomPaperHeight });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000(object background)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", background);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void PrintOut2000(object background, object append)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", background, append);
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
		public virtual void PrintOut2000(object background, object append, object range)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", background, append, range);
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
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", background, append, range, outputFileName);
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
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from });
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
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to });
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
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item });
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
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies });
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
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages });
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
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType });
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
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile });
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
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate });
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
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX });
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
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint });
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
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn });
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
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow });
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
		public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838511.aspx </remarks>
		/// <param name="codePageOrigin">Int32 codePageOrigin</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void ConvertVietDoc(Int32 codePageOrigin)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ConvertVietDoc", codePageOrigin);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194029.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		/// <param name="makePublic">optional bool MakePublic = false</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void CheckIn(object saveChanges, object comments, object makePublic)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckIn", saveChanges, comments, makePublic);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194029.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void CheckIn()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckIn");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194029.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void CheckIn(object saveChanges)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckIn", saveChanges);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194029.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void CheckIn(object saveChanges, object comments)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckIn", saveChanges, comments);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198206.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual bool CanCheckin()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "CanCheckin");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193054.aspx </remarks>
		/// <param name="recipients">optional object recipients</param>
		/// <param name="subject">optional object subject</param>
		/// <param name="showMessage">optional object showMessage</param>
		/// <param name="includeAttachment">optional object includeAttachment</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void SendForReview(object recipients, object subject, object showMessage, object includeAttachment)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendForReview", recipients, subject, showMessage, includeAttachment);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193054.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void SendForReview()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendForReview");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193054.aspx </remarks>
		/// <param name="recipients">optional object recipients</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void SendForReview(object recipients)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendForReview", recipients);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193054.aspx </remarks>
		/// <param name="recipients">optional object recipients</param>
		/// <param name="subject">optional object subject</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void SendForReview(object recipients, object subject)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendForReview", recipients, subject);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193054.aspx </remarks>
		/// <param name="recipients">optional object recipients</param>
		/// <param name="subject">optional object subject</param>
		/// <param name="showMessage">optional object showMessage</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void SendForReview(object recipients, object subject, object showMessage)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendForReview", recipients, subject, showMessage);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836324.aspx </remarks>
		/// <param name="showMessage">optional object showMessage</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void ReplyWithChanges(object showMessage)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ReplyWithChanges", showMessage);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836324.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void ReplyWithChanges()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ReplyWithChanges");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837660.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void EndReview()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EndReview");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195460.aspx </remarks>
		/// <param name="passwordEncryptionProvider">string passwordEncryptionProvider</param>
		/// <param name="passwordEncryptionAlgorithm">string passwordEncryptionAlgorithm</param>
		/// <param name="passwordEncryptionKeyLength">Int32 passwordEncryptionKeyLength</param>
		/// <param name="passwordEncryptionFileProperties">optional object passwordEncryptionFileProperties</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void SetPasswordEncryptionOptions(string passwordEncryptionProvider, string passwordEncryptionAlgorithm, Int32 passwordEncryptionKeyLength, object passwordEncryptionFileProperties)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetPasswordEncryptionOptions", passwordEncryptionProvider, passwordEncryptionAlgorithm, passwordEncryptionKeyLength, passwordEncryptionFileProperties);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195460.aspx </remarks>
		/// <param name="passwordEncryptionProvider">string passwordEncryptionProvider</param>
		/// <param name="passwordEncryptionAlgorithm">string passwordEncryptionAlgorithm</param>
		/// <param name="passwordEncryptionKeyLength">Int32 passwordEncryptionKeyLength</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void SetPasswordEncryptionOptions(string passwordEncryptionProvider, string passwordEncryptionAlgorithm, Int32 passwordEncryptionKeyLength)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetPasswordEncryptionOptions", passwordEncryptionProvider, passwordEncryptionAlgorithm, passwordEncryptionKeyLength);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void RecheckSmartTags()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RecheckSmartTags");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void RemoveSmartTags()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RemoveSmartTags");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198118.aspx </remarks>
		/// <param name="style">object style</param>
		/// <param name="setInTemplate">bool setInTemplate</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void SetDefaultTableStyle(object style, bool setInTemplate)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetDefaultTableStyle", style, setInTemplate);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822910.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void DeleteAllComments()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteAllComments");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837501.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void AcceptAllRevisionsShown()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AcceptAllRevisionsShown");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822533.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void RejectAllRevisionsShown()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RejectAllRevisionsShown");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836620.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void DeleteAllCommentsShown()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteAllCommentsShown");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821137.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void ResetFormFields()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ResetFormFields");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void CheckNewSmartTags()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckNewSmartTags");
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType type</param>
		/// <param name="noReset">optional object noReset</param>
		/// <param name="password">optional object password</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void Protect2002(NetOffice.WordApi.Enums.WdProtectionType type, object noReset, object password)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Protect2002", type, noReset, password);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType type</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void Protect2002(NetOffice.WordApi.Enums.WdProtectionType type)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Protect2002", type);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.WordApi.Enums.WdProtectionType type</param>
		/// <param name="noReset">optional object noReset</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void Protect2002(NetOffice.WordApi.Enums.WdProtectionType type, object noReset)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Protect2002", type, noReset);
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
		public virtual void Compare2002(string name, object authorName, object compareTarget, object detectFormatChanges, object ignoreAllComparisonWarnings, object addToRecentFiles)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Compare2002", new object[]{ name, authorName, compareTarget, detectFormatChanges, ignoreAllComparisonWarnings, addToRecentFiles });
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void Compare2002(string name)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Compare2002", name);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="authorName">optional object authorName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void Compare2002(string name, object authorName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Compare2002", name, authorName);
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
		public virtual void Compare2002(string name, object authorName, object compareTarget)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Compare2002", name, authorName, compareTarget);
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
		public virtual void Compare2002(string name, object authorName, object compareTarget, object detectFormatChanges)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Compare2002", name, authorName, compareTarget, detectFormatChanges);
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
		public virtual void Compare2002(string name, object authorName, object compareTarget, object detectFormatChanges, object ignoreAllComparisonWarnings)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Compare2002", new object[]{ name, authorName, compareTarget, detectFormatChanges, ignoreAllComparisonWarnings });
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192018.aspx </remarks>
		/// <param name="recipients">optional object recipients</param>
		/// <param name="subject">optional object subject</param>
		/// <param name="showMessage">optional object showMessage</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void SendFaxOverInternet(object recipients, object subject, object showMessage)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendFaxOverInternet", recipients, subject, showMessage);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192018.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void SendFaxOverInternet()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendFaxOverInternet");
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192018.aspx </remarks>
		/// <param name="recipients">optional object recipients</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void SendFaxOverInternet(object recipients)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendFaxOverInternet", recipients);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192018.aspx </remarks>
		/// <param name="recipients">optional object recipients</param>
		/// <param name="subject">optional object subject</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void SendFaxOverInternet(object recipients, object subject)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendFaxOverInternet", recipients, subject);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196274.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="dataOnly">optional bool DataOnly = true</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void TransformDocument(string path, object dataOnly)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransformDocument", path, dataOnly);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196274.aspx </remarks>
		/// <param name="path">string path</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void TransformDocument(string path)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TransformDocument", path);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195660.aspx </remarks>
		/// <param name="editorID">optional object editorID</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void SelectAllEditableRanges(object editorID)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SelectAllEditableRanges", editorID);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195660.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void SelectAllEditableRanges()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SelectAllEditableRanges");
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844883.aspx </remarks>
		/// <param name="editorID">optional object editorID</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void DeleteAllEditableRanges(object editorID)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteAllEditableRanges", editorID);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844883.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void DeleteAllEditableRanges()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteAllEditableRanges");
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838947.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void DeleteAllInkAnnotations()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteAllInkAnnotations");
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
		public virtual void AddDocumentWorkspaceHeader(bool richFormat, string url, string title, string description, string iD)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddDocumentWorkspaceHeader", new object[]{ richFormat, url, title, description, iD });
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="iD">string iD</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void RemoveDocumentWorkspaceHeader(string iD)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RemoveDocumentWorkspaceHeader", iD);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845389.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void RemoveLockedStyles()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RemoveLockedStyles");
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822346.aspx </remarks>
		/// <param name="xPath">string xPath</param>
		/// <param name="prefixMapping">optional string PrefixMapping = </param>
		/// <param name="fastSearchSkippingTextNodes">optional bool FastSearchSkippingTextNodes = true</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual NetOffice.WordApi.XMLNode SelectSingleNode(string xPath, object prefixMapping, object fastSearchSkippingTextNodes)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLNode>(this, "SelectSingleNode", typeof(NetOffice.WordApi.XMLNode), xPath, prefixMapping, fastSearchSkippingTextNodes);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822346.aspx </remarks>
		/// <param name="xPath">string xPath</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual NetOffice.WordApi.XMLNode SelectSingleNode(string xPath)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLNode>(this, "SelectSingleNode", typeof(NetOffice.WordApi.XMLNode), xPath);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822346.aspx </remarks>
		/// <param name="xPath">string xPath</param>
		/// <param name="prefixMapping">optional string PrefixMapping = </param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual NetOffice.WordApi.XMLNode SelectSingleNode(string xPath, object prefixMapping)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLNode>(this, "SelectSingleNode", typeof(NetOffice.WordApi.XMLNode), xPath, prefixMapping);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837252.aspx </remarks>
		/// <param name="xPath">string xPath</param>
		/// <param name="prefixMapping">optional string PrefixMapping = </param>
		/// <param name="fastSearchSkippingTextNodes">optional bool FastSearchSkippingTextNodes = true</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual NetOffice.WordApi.XMLNodes SelectNodes(string xPath, object prefixMapping, object fastSearchSkippingTextNodes)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLNodes>(this, "SelectNodes", typeof(NetOffice.WordApi.XMLNodes), xPath, prefixMapping, fastSearchSkippingTextNodes);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837252.aspx </remarks>
		/// <param name="xPath">string xPath</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual NetOffice.WordApi.XMLNodes SelectNodes(string xPath)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLNodes>(this, "SelectNodes", typeof(NetOffice.WordApi.XMLNodes), xPath);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837252.aspx </remarks>
		/// <param name="xPath">string xPath</param>
		/// <param name="prefixMapping">optional string PrefixMapping = </param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual NetOffice.WordApi.XMLNodes SelectNodes(string xPath, object prefixMapping)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLNodes>(this, "SelectNodes", typeof(NetOffice.WordApi.XMLNodes), xPath, prefixMapping);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197270.aspx </remarks>
		/// <param name="removeDocInfoType">NetOffice.WordApi.Enums.WdRemoveDocInfoType removeDocInfoType</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void RemoveDocumentInformation(NetOffice.WordApi.Enums.WdRemoveDocInfoType removeDocInfoType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RemoveDocumentInformation", removeDocInfoType);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840884.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		/// <param name="makePublic">optional bool MakePublic = false</param>
		/// <param name="versionType">optional object versionType</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void CheckInWithVersion(object saveChanges, object comments, object makePublic, object versionType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckInWithVersion", saveChanges, comments, makePublic, versionType);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840884.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void CheckInWithVersion()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckInWithVersion");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840884.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void CheckInWithVersion(object saveChanges)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckInWithVersion", saveChanges);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840884.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void CheckInWithVersion(object saveChanges, object comments)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckInWithVersion", saveChanges, comments);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840884.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		/// <param name="makePublic">optional bool MakePublic = false</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void CheckInWithVersion(object saveChanges, object comments, object makePublic)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckInWithVersion", saveChanges, comments, makePublic);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void Dummy2()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Dummy2");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845518.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void LockServerFile()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LockServerFile");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198071.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.OfficeApi.WorkflowTasks GetWorkflowTasks()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.WorkflowTasks>(this, "GetWorkflowTasks", typeof(NetOffice.OfficeApi.WorkflowTasks));
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845242.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.OfficeApi.WorkflowTemplates GetWorkflowTemplates()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.WorkflowTemplates>(this, "GetWorkflowTemplates", typeof(NetOffice.OfficeApi.WorkflowTemplates));
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void Dummy4()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Dummy4");
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
		public virtual void AddMeetingWorkspaceHeader(bool skipIfAbsent, string url, string title, string description, string iD)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddMeetingWorkspaceHeader", new object[]{ skipIfAbsent, url, title, description, iD });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198291.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void SaveAsQuickStyleSet(string fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAsQuickStyleSet", fileName);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void ApplyQuickStyleSet(string name)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyQuickStyleSet", name);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840910.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void ApplyDocumentTheme(string fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDocumentTheme", fileName);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838276.aspx </remarks>
		/// <param name="node">NetOffice.OfficeApi.CustomXMLNode node</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.ContentControls SelectLinkedControls(NetOffice.OfficeApi.CustomXMLNode node)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.ContentControls>(this, "SelectLinkedControls", typeof(NetOffice.WordApi.ContentControls), node);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198010.aspx </remarks>
		/// <param name="stream">optional NetOffice.OfficeApi.CustomXMLPart Stream = 0</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.ContentControls SelectUnlinkedControls(object stream)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.ContentControls>(this, "SelectUnlinkedControls", typeof(NetOffice.WordApi.ContentControls), stream);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198010.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.ContentControls SelectUnlinkedControls()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.ContentControls>(this, "SelectUnlinkedControls", typeof(NetOffice.WordApi.ContentControls));
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822990.aspx </remarks>
		/// <param name="title">string title</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.ContentControls SelectContentControlsByTitle(string title)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.ContentControls>(this, "SelectContentControlsByTitle", typeof(NetOffice.WordApi.ContentControls), title);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
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
		public virtual void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags, object bitmapMissingFonts, object useISO19005_1, object fixedFormatExtClassPtr)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, range, from, to, item, includeDocProps, keepIRM, createBookmarks, docStructureTags, bitmapMissingFonts, useISO19005_1, fixedFormatExtClassPtr });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", outputFileName, exportFormat);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", outputFileName, exportFormat, openAfterExport);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", outputFileName, exportFormat, openAfterExport, optimizeFor);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, range });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, range, from });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Enums.WdExportRange Range = 0</param>
		/// <param name="from">optional Int32 From = 1</param>
		/// <param name="to">optional Int32 To = 1</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, range, from, to });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
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
		public virtual void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, range, from, to, item });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
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
		public virtual void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, range, from, to, item, includeDocProps });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
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
		public virtual void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps, object keepIRM)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, range, from, to, item, includeDocProps, keepIRM });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
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
		public virtual void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps, object keepIRM, object createBookmarks)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, range, from, to, item, includeDocProps, keepIRM, createBookmarks });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
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
		public virtual void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, range, from, to, item, includeDocProps, keepIRM, createBookmarks, docStructureTags });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
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
		public virtual void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags, object bitmapMissingFonts)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, range, from, to, item, includeDocProps, keepIRM, createBookmarks, docStructureTags, bitmapMissingFonts });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840962.aspx </remarks>
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
		public virtual void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object range, object from, object to, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags, object bitmapMissingFonts, object useISO19005_1)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, range, from, to, item, includeDocProps, keepIRM, createBookmarks, docStructureTags, bitmapMissingFonts, useISO19005_1 });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196504.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void FreezeLayout()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FreezeLayout");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void UnfreezeLayout()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "UnfreezeLayout");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194276.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void DowngradeDocument()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DowngradeDocument");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835714.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void Convert()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Convert");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839693.aspx </remarks>
		/// <param name="tag">string tag</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.ContentControls SelectContentControlsByTag(string tag)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.ContentControls>(this, "SelectContentControlsByTag", typeof(NetOffice.WordApi.ContentControls), tag);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838360.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public virtual void ConvertAutoHyphens()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ConvertAutoHyphens");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821672.aspx </remarks>
		/// <param name="style">object style</param>
		[SupportByVersion("Word", 14,15,16)]
		public virtual void ApplyQuickStyleSet2(object style)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyQuickStyleSet2", style);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
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
		public virtual void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding, object addBiDiMarks, object compatibilityMode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions, lineEnding, addBiDiMarks, compatibilityMode });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public virtual void SaveAs2()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
		/// <param name="fileName">optional object fileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public virtual void SaveAs2(object fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2", fileName);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public virtual void SaveAs2(object fileName, object fileFormat)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2", fileName, fileFormat);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public virtual void SaveAs2(object fileName, object fileFormat, object lockComments)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2", fileName, fileFormat, lockComments);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public virtual void SaveAs2(object fileName, object fileFormat, object lockComments, object password)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2", fileName, fileFormat, lockComments, password);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public virtual void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public virtual void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public virtual void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
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
		public virtual void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
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
		public virtual void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
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
		public virtual void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
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
		public virtual void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
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
		public virtual void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
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
		public virtual void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
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
		public virtual void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
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
		public virtual void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions, lineEnding });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836084.aspx </remarks>
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
		public virtual void SaveAs2(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding, object addBiDiMarks)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs2", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions, lineEnding, addBiDiMarks });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840359.aspx </remarks>
		/// <param name="mode">Int32 mode</param>
		[SupportByVersion("Word", 14,15,16)]
		public virtual void SetCompatibilityMode(Int32 mode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetCompatibilityMode", mode);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231927.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public virtual Int32 ReturnToLastReadPosition()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ReturnToLastReadPosition");
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
		public virtual void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding, object addBiDiMarks, object compatibilityMode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveCopyAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions, lineEnding, addBiDiMarks, compatibilityMode });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public virtual void SaveCopyAs()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveCopyAs");
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public virtual void SaveCopyAs(object fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveCopyAs", fileName);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public virtual void SaveCopyAs(object fileName, object fileFormat)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveCopyAs", fileName, fileFormat);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="lockComments">optional object lockComments</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public virtual void SaveCopyAs(object fileName, object fileFormat, object lockComments)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveCopyAs", fileName, fileFormat, lockComments);
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
		public virtual void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveCopyAs", fileName, fileFormat, lockComments, password);
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
		public virtual void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveCopyAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles });
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
		public virtual void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveCopyAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword });
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
		public virtual void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveCopyAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended });
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
		public virtual void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveCopyAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts });
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
		public virtual void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveCopyAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat });
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
		public virtual void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveCopyAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData });
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
		public virtual void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveCopyAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter });
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
		public virtual void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveCopyAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding });
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
		public virtual void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveCopyAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks });
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
		public virtual void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveCopyAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions });
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
		public virtual void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveCopyAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions, lineEnding });
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
		public virtual void SaveCopyAs(object fileName, object fileFormat, object lockComments, object password, object addToRecentFiles, object writePassword, object readOnlyRecommended, object embedTrueTypeFonts, object saveNativePictureFormat, object saveFormsData, object saveAsAOCELetter, object encoding, object insertLineBreaks, object allowSubstitutions, object lineEnding, object addBiDiMarks)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveCopyAs", new object[]{ fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts, saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions, lineEnding, addBiDiMarks });
		}

		#endregion

		#pragma warning restore
	}
}


