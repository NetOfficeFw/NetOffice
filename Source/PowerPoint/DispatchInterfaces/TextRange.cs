using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface TextRange 
	/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745366.aspx </remarks>
	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class TextRange : Collection
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744897.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Application>(this, "Application", NetOffice.PowerPointApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744386.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745346.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.ActionSettings ActionSettings
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ActionSettings>(this, "ActionSettings", NetOffice.PowerPointApi.ActionSettings.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744180.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Int32 Start
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Start");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744845.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Int32 Length
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Length");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744266.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Single BoundLeft
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "BoundLeft");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746319.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Single BoundTop
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "BoundTop");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744683.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Single BoundWidth
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "BoundWidth");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745606.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Single BoundHeight
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "BoundHeight");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746239.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744240.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Font Font
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Font>(this, "Font", NetOffice.PowerPointApi.Font.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744700.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.ParagraphFormat ParagraphFormat
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ParagraphFormat>(this, "ParagraphFormat", NetOffice.PowerPointApi.ParagraphFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744603.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Int32 IndentLevel
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "IndentLevel");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IndentLevel", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746734.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744863.aspx </remarks>
		/// <param name="start">optional Int32 Start = -1</param>
		/// <param name="length">optional Int32 Length = -1</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange Paragraphs(object start, object length)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "Paragraphs", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, start, length);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744863.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange Paragraphs()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "Paragraphs", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744863.aspx </remarks>
		/// <param name="start">optional Int32 Start = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange Paragraphs(object start)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "Paragraphs", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, start);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746189.aspx </remarks>
		/// <param name="start">optional Int32 Start = -1</param>
		/// <param name="length">optional Int32 Length = -1</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange Sentences(object start, object length)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "Sentences", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, start, length);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746189.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange Sentences()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "Sentences", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746189.aspx </remarks>
		/// <param name="start">optional Int32 Start = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange Sentences(object start)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "Sentences", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, start);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746049.aspx </remarks>
		/// <param name="start">optional Int32 Start = -1</param>
		/// <param name="length">optional Int32 Length = -1</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange Words(object start, object length)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "Words", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, start, length);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746049.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange Words()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "Words", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746049.aspx </remarks>
		/// <param name="start">optional Int32 Start = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange Words(object start)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "Words", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, start);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743845.aspx </remarks>
		/// <param name="start">optional Int32 Start = -1</param>
		/// <param name="length">optional Int32 Length = -1</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange Characters(object start, object length)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "Characters", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, start, length);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743845.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange Characters()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "Characters", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743845.aspx </remarks>
		/// <param name="start">optional Int32 Start = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange Characters(object start)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "Characters", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, start);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745593.aspx </remarks>
		/// <param name="start">optional Int32 Start = -1</param>
		/// <param name="length">optional Int32 Length = -1</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange Lines(object start, object length)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "Lines", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, start, length);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745593.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange Lines()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "Lines", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745593.aspx </remarks>
		/// <param name="start">optional Int32 Start = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange Lines(object start)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "Lines", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, start);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743973.aspx </remarks>
		/// <param name="start">optional Int32 Start = -1</param>
		/// <param name="length">optional Int32 Length = -1</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange Runs(object start, object length)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "Runs", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, start, length);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743973.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange Runs()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "Runs", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743973.aspx </remarks>
		/// <param name="start">optional Int32 Start = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange Runs(object start)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "Runs", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, start);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745464.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange TrimText()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "TrimText", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744323.aspx </remarks>
		/// <param name="newText">optional string NewText = </param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange InsertAfter(object newText)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "InsertAfter", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, newText);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744323.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange InsertAfter()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "InsertAfter", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746793.aspx </remarks>
		/// <param name="newText">optional string NewText = </param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange InsertBefore(object newText)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "InsertBefore", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, newText);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746793.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange InsertBefore()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "InsertBefore", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745977.aspx </remarks>
		/// <param name="dateTimeFormat">NetOffice.PowerPointApi.Enums.PpDateTimeFormat dateTimeFormat</param>
		/// <param name="insertAsField">optional NetOffice.OfficeApi.Enums.MsoTriState InsertAsField = 0</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange InsertDateTime(NetOffice.PowerPointApi.Enums.PpDateTimeFormat dateTimeFormat, object insertAsField)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "InsertDateTime", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, dateTimeFormat, insertAsField);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745977.aspx </remarks>
		/// <param name="dateTimeFormat">NetOffice.PowerPointApi.Enums.PpDateTimeFormat dateTimeFormat</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange InsertDateTime(NetOffice.PowerPointApi.Enums.PpDateTimeFormat dateTimeFormat)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "InsertDateTime", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, dateTimeFormat);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743923.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange InsertSlideNumber()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "InsertSlideNumber", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745835.aspx </remarks>
		/// <param name="fontName">string fontName</param>
		/// <param name="charNumber">Int32 charNumber</param>
		/// <param name="unicode">optional NetOffice.OfficeApi.Enums.MsoTriState Unicode = 0</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange InsertSymbol(string fontName, Int32 charNumber, object unicode)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "InsertSymbol", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, fontName, charNumber, unicode);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745835.aspx </remarks>
		/// <param name="fontName">string fontName</param>
		/// <param name="charNumber">Int32 charNumber</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange InsertSymbol(string fontName, Int32 charNumber)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "InsertSymbol", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, fontName, charNumber);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746287.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Select()
		{
			 Factory.ExecuteMethod(this, "Select");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745756.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Cut()
		{
			 Factory.ExecuteMethod(this, "Cut");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746251.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Copy()
		{
			 Factory.ExecuteMethod(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744334.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744812.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange Paste()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "Paste", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745812.aspx </remarks>
		/// <param name="type">NetOffice.PowerPointApi.Enums.PpChangeCase type</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void ChangeCase(NetOffice.PowerPointApi.Enums.PpChangeCase type)
		{
			 Factory.ExecuteMethod(this, "ChangeCase", type);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744944.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void AddPeriods()
		{
			 Factory.ExecuteMethod(this, "AddPeriods");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745113.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void RemovePeriods()
		{
			 Factory.ExecuteMethod(this, "RemovePeriods");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744247.aspx </remarks>
		/// <param name="findWhat">string findWhat</param>
		/// <param name="after">optional Int32 After = 0</param>
		/// <param name="matchCase">optional NetOffice.OfficeApi.Enums.MsoTriState MatchCase = 0</param>
		/// <param name="wholeWords">optional NetOffice.OfficeApi.Enums.MsoTriState WholeWords = 0</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange Find(string findWhat, object after, object matchCase, object wholeWords)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "Find", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, findWhat, after, matchCase, wholeWords);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744247.aspx </remarks>
		/// <param name="findWhat">string findWhat</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange Find(string findWhat)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "Find", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, findWhat);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744247.aspx </remarks>
		/// <param name="findWhat">string findWhat</param>
		/// <param name="after">optional Int32 After = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange Find(string findWhat, object after)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "Find", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, findWhat, after);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744247.aspx </remarks>
		/// <param name="findWhat">string findWhat</param>
		/// <param name="after">optional Int32 After = 0</param>
		/// <param name="matchCase">optional NetOffice.OfficeApi.Enums.MsoTriState MatchCase = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange Find(string findWhat, object after, object matchCase)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "Find", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, findWhat, after, matchCase);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743889.aspx </remarks>
		/// <param name="findWhat">string findWhat</param>
		/// <param name="replaceWhat">string replaceWhat</param>
		/// <param name="after">optional Int32 After = 0</param>
		/// <param name="matchCase">optional NetOffice.OfficeApi.Enums.MsoTriState MatchCase = 0</param>
		/// <param name="wholeWords">optional NetOffice.OfficeApi.Enums.MsoTriState WholeWords = 0</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange Replace(string findWhat, string replaceWhat, object after, object matchCase, object wholeWords)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "Replace", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, new object[]{ findWhat, replaceWhat, after, matchCase, wholeWords });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743889.aspx </remarks>
		/// <param name="findWhat">string findWhat</param>
		/// <param name="replaceWhat">string replaceWhat</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange Replace(string findWhat, string replaceWhat)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "Replace", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, findWhat, replaceWhat);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743889.aspx </remarks>
		/// <param name="findWhat">string findWhat</param>
		/// <param name="replaceWhat">string replaceWhat</param>
		/// <param name="after">optional Int32 After = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange Replace(string findWhat, string replaceWhat, object after)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "Replace", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, findWhat, replaceWhat, after);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743889.aspx </remarks>
		/// <param name="findWhat">string findWhat</param>
		/// <param name="replaceWhat">string replaceWhat</param>
		/// <param name="after">optional Int32 After = 0</param>
		/// <param name="matchCase">optional NetOffice.OfficeApi.Enums.MsoTriState MatchCase = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange Replace(string findWhat, string replaceWhat, object after, object matchCase)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "Replace", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, findWhat, replaceWhat, after, matchCase);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744523.aspx </remarks>
		/// <param name="x1">Single x1</param>
		/// <param name="y1">Single y1</param>
		/// <param name="x2">Single x2</param>
		/// <param name="y2">Single y2</param>
		/// <param name="x3">Single x3</param>
		/// <param name="y3">Single y3</param>
		/// <param name="x4">Single x4</param>
		/// <param name="y4">Single y4</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void RotatedBounds(out Single x1, out Single y1, out Single x2, out Single y2, out Single x3, out Single y3, out Single x4, out Single y4)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,true,true,true,true,true,true,true);
			x1 = 0;
			y1 = 0;
			x2 = 0;
			y2 = 0;
			x3 = 0;
			y3 = 0;
			x4 = 0;
			y4 = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(x1, y1, x2, y2, x3, y3, x4, y4);
			Invoker.Method(this, "RotatedBounds", paramsArray, modifiers);
			x1 = (Single)paramsArray[0];
			y1 = (Single)paramsArray[1];
			x2 = (Single)paramsArray[2];
			y2 = (Single)paramsArray[3];
			x3 = (Single)paramsArray[4];
			y3 = (Single)paramsArray[5];
			x4 = (Single)paramsArray[6];
			y4 = (Single)paramsArray[7];
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746623.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void RtlRun()
		{
			 Factory.ExecuteMethod(this, "RtlRun");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744980.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void LtrRun()
		{
			 Factory.ExecuteMethod(this, "LtrRun");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745706.aspx </remarks>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		/// <param name="iconFileName">optional string IconFileName = </param>
		/// <param name="iconIndex">optional Int32 IconIndex = 0</param>
		/// <param name="iconLabel">optional string IconLabel = </param>
		/// <param name="link">optional NetOffice.OfficeApi.Enums.MsoTriState Link = 0</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange PasteSpecial(object dataType, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel, object link)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "PasteSpecial", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, new object[]{ dataType, displayAsIcon, iconFileName, iconIndex, iconLabel, link });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745706.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange PasteSpecial()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "PasteSpecial", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745706.aspx </remarks>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange PasteSpecial(object dataType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "PasteSpecial", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, dataType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745706.aspx </remarks>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange PasteSpecial(object dataType, object displayAsIcon)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "PasteSpecial", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, dataType, displayAsIcon);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745706.aspx </remarks>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		/// <param name="iconFileName">optional string IconFileName = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange PasteSpecial(object dataType, object displayAsIcon, object iconFileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "PasteSpecial", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, dataType, displayAsIcon, iconFileName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745706.aspx </remarks>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		/// <param name="iconFileName">optional string IconFileName = </param>
		/// <param name="iconIndex">optional Int32 IconIndex = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange PasteSpecial(object dataType, object displayAsIcon, object iconFileName, object iconIndex)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "PasteSpecial", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, dataType, displayAsIcon, iconFileName, iconIndex);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745706.aspx </remarks>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		/// <param name="iconFileName">optional string IconFileName = </param>
		/// <param name="iconIndex">optional Int32 IconIndex = 0</param>
		/// <param name="iconLabel">optional string IconLabel = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.TextRange PasteSpecial(object dataType, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.TextRange>(this, "PasteSpecial", NetOffice.PowerPointApi.TextRange.LateBindingApiWrapperType, new object[]{ dataType, displayAsIcon, iconFileName, iconIndex, iconLabel });
		}

		#endregion

		#pragma warning restore
	}
}
