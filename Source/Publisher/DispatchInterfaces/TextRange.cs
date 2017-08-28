using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface TextRange 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class TextRange : COMObject
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
                    _type = typeof(TextRange);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public TextRange(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public TextRange(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextRange(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextRange(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextRange(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextRange(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextRange() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextRange(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public string Text
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text", value);
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
		public NetOffice.PublisherApi.Font Font
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Font>(this, "Font", NetOffice.PublisherApi.Font.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Font", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 Length
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Length");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 Start
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Start");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Start", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 End
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "End");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "End", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single BoundLeft
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "BoundLeft");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single BoundHeight
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "BoundHeight");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single BoundTop
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "BoundTop");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single BoundWidth
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "BoundWidth");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ParagraphFormat ParagraphFormat
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ParagraphFormat>(this, "ParagraphFormat", NetOffice.PublisherApi.ParagraphFormat.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "ParagraphFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		public object ContainingObject
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "ContainingObject");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange Duplicate
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.TextRange>(this, "Duplicate", NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Font MajorityFont
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Font>(this, "MajorityFont", NetOffice.PublisherApi.Font.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ParagraphFormat MajorityParagraphFormat
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ParagraphFormat>(this, "MajorityParagraphFormat", NetOffice.PublisherApi.ParagraphFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Fields Fields
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Fields>(this, "Fields", NetOffice.PublisherApi.Fields.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Story Story
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Story>(this, "Story", NetOffice.PublisherApi.Story.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoLanguageID LanguageID
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoLanguageID>(this, "LanguageID");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "LanguageID", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.DropCap DropCap
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.DropCap>(this, "DropCap", NetOffice.PublisherApi.DropCap.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbFontScriptType Script
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbFontScriptType>(this, "Script");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Hyperlinks Hyperlinks
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Hyperlinks>(this, "Hyperlinks", NetOffice.PublisherApi.Hyperlinks.LateBindingApiWrapperType);
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
		public NetOffice.PublisherApi.InlineShapes InlineShapes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.InlineShapes>(this, "InlineShapes", NetOffice.PublisherApi.InlineShapes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 WordsCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "WordsCount");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 LinesCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "LinesCount");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 ParagraphsCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ParagraphsCount");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="direction">NetOffice.PublisherApi.Enums.PbCollapseDirection direction</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void Collapse(NetOffice.PublisherApi.Enums.PbCollapseDirection direction)
		{
			 Factory.ExecuteMethod(this, "Collapse", direction);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="unit">NetOffice.PublisherApi.Enums.PbTextUnit unit</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 Expand(NetOffice.PublisherApi.Enums.PbTextUnit unit)
		{
			return Factory.ExecuteInt32MethodGet(this, "Expand", unit);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="unit">NetOffice.PublisherApi.Enums.PbTextUnit unit</param>
		/// <param name="size">Int32 size</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 Move(NetOffice.PublisherApi.Enums.PbTextUnit unit, Int32 size)
		{
			return Factory.ExecuteInt32MethodGet(this, "Move", unit, size);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="unit">NetOffice.PublisherApi.Enums.PbTextUnit unit</param>
		/// <param name="size">Int32 size</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 MoveStart(NetOffice.PublisherApi.Enums.PbTextUnit unit, Int32 size)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveStart", unit, size);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="unit">NetOffice.PublisherApi.Enums.PbTextUnit unit</param>
		/// <param name="size">Int32 size</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 MoveEnd(NetOffice.PublisherApi.Enums.PbTextUnit unit, Int32 size)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveEnd", unit, size);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="start">Int32 start</param>
		/// <param name="length">optional Int32 Length = 1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange Characters(Int32 start, object length)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.TextRange>(this, "Characters", NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType, start, length);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="start">Int32 start</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange Characters(Int32 start)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.TextRange>(this, "Characters", NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType, start);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="newText">string newText</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange InsertAfter(string newText)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.TextRange>(this, "InsertAfter", NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType, newText);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="newText">string newText</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange InsertBefore(string newText)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.TextRange>(this, "InsertBefore", NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType, newText);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="fontName">string fontName</param>
		/// <param name="charIndex">Int32 charIndex</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange InsertSymbol(string fontName, Int32 charIndex)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.TextRange>(this, "InsertSymbol", NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType, fontName, charIndex);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbDateTimeFormat format</param>
		/// <param name="insertAsField">optional bool InsertAsField = false</param>
		/// <param name="insertAsFullWidth">optional bool InsertAsFullWidth = false</param>
		/// <param name="language">optional NetOffice.OfficeApi.Enums.MsoLanguageID Language = 0</param>
		/// <param name="calendar">optional NetOffice.PublisherApi.Enums.PbCalendarType Calendar = 0</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange InsertDateTime(NetOffice.PublisherApi.Enums.PbDateTimeFormat format, object insertAsField, object insertAsFullWidth, object language, object calendar)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.TextRange>(this, "InsertDateTime", NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType, new object[]{ format, insertAsField, insertAsFullWidth, language, calendar });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbDateTimeFormat format</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange InsertDateTime(NetOffice.PublisherApi.Enums.PbDateTimeFormat format)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.TextRange>(this, "InsertDateTime", NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType, format);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbDateTimeFormat format</param>
		/// <param name="insertAsField">optional bool InsertAsField = false</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange InsertDateTime(NetOffice.PublisherApi.Enums.PbDateTimeFormat format, object insertAsField)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.TextRange>(this, "InsertDateTime", NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType, format, insertAsField);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbDateTimeFormat format</param>
		/// <param name="insertAsField">optional bool InsertAsField = false</param>
		/// <param name="insertAsFullWidth">optional bool InsertAsFullWidth = false</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange InsertDateTime(NetOffice.PublisherApi.Enums.PbDateTimeFormat format, object insertAsField, object insertAsFullWidth)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.TextRange>(this, "InsertDateTime", NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType, format, insertAsField, insertAsFullWidth);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbDateTimeFormat format</param>
		/// <param name="insertAsField">optional bool InsertAsField = false</param>
		/// <param name="insertAsFullWidth">optional bool InsertAsFullWidth = false</param>
		/// <param name="language">optional NetOffice.OfficeApi.Enums.MsoLanguageID Language = 0</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange InsertDateTime(NetOffice.PublisherApi.Enums.PbDateTimeFormat format, object insertAsField, object insertAsFullWidth, object language)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.TextRange>(this, "InsertDateTime", NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType, format, insertAsField, insertAsFullWidth, language);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="start">Int32 start</param>
		/// <param name="length">optional Int32 Length = 1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange Paragraphs(Int32 start, object length)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.TextRange>(this, "Paragraphs", NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType, start, length);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="start">Int32 start</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange Paragraphs(Int32 start)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.TextRange>(this, "Paragraphs", NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType, start);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="start">Int32 start</param>
		/// <param name="length">optional Int32 Length = 1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange Lines(Int32 start, object length)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.TextRange>(this, "Lines", NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType, start, length);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="start">Int32 start</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange Lines(Int32 start)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.TextRange>(this, "Lines", NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType, start);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="start">Int32 start</param>
		/// <param name="length">optional Int32 Length = 1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange Words(Int32 start, object length)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.TextRange>(this, "Words", NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType, start, length);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="start">Int32 start</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange Words(Int32 start)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.TextRange>(this, "Words", NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType, start);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void Select()
		{
			 Factory.ExecuteMethod(this, "Select");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void Cut()
		{
			 Factory.ExecuteMethod(this, "Cut");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void Copy()
		{
			 Factory.ExecuteMethod(this, "Copy");
		}

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
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange Paste()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.TextRange>(this, "Paste", NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="varIndex">object varIndex</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange InsertMailMergeField(object varIndex)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.TextRange>(this, "InsertMailMergeField", NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType, varIndex);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.PublisherApi.Enums.PbPageNumberType Type = 1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange InsertPageNumber(object type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.TextRange>(this, "InsertPageNumber", NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType, type);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange InsertPageNumber()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.TextRange>(this, "InsertPageNumber", NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange InsertBarcode()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.TextRange>(this, "InsertBarcode", NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType);
		}

		#endregion

		#pragma warning restore
	}
}
